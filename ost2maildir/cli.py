"""CLI entry point for ost2maildir."""

import datetime
import logging
import os
import sys
from pathlib import Path

import click
from tqdm import tqdm

from .calendar_conv import convert_calendar
from .contacts_conv import convert_contact
from .converter import MIMEConverter
from .dedup import Deduplicator
from .folders import FolderMapper, SKIP_FOLDER_NAMES
from .maildir import MaildirWriter
from .reader import OSTReader, EncryptionError
from .utils import clean_subject, entry_as_string, get_all_mapi_props, synthetic_message_id

log = logging.getLogger("ost2maildir")

PR_MESSAGE_CLASS = 0x001A
PR_INTERNET_MESSAGE_ID = 0x1035
PR_MESSAGE_DELIVERY_TIME = 0x0E06
PR_CLIENT_SUBMIT_TIME = 0x0039
PR_CREATION_TIME = 0x3007


@click.command()
@click.argument("input_ost", type=click.Path(exists=True, readable=True, path_type=Path))
@click.argument("output_maildir", type=click.Path(path_type=Path))
@click.option("--no-dedup", is_flag=True, help="Disable Message-ID deduplication.")
@click.option("--skip-contacts", is_flag=True, help="Do not convert contacts.")
@click.option("--skip-calendar", is_flag=True, help="Do not convert calendar items.")
@click.option("--no-tnef", is_flag=True, help="Preserve winmail.dat without decoding.")
@click.option("--encoding", default="cp1252", show_default=True, help="Fallback encoding for ANSI bodies.")
@click.option("--log-file", type=click.Path(path_type=Path), help="Write log to file.")
@click.option("--dry-run", is_flag=True, help="Count items without writing output.")
@click.option("--hostname", default=None, help="Hostname for synthetic Message-IDs.")
@click.option("-v", "--verbose", is_flag=True, help="Debug-level logging.")
@click.version_option("1.0.0")
def main(
    input_ost: Path,
    output_maildir: Path,
    no_dedup: bool,
    skip_contacts: bool,
    skip_calendar: bool,
    no_tnef: bool,
    encoding: str,
    log_file: Path | None,
    dry_run: bool,
    hostname: str | None,
    verbose: bool,
):
    """Convert a Microsoft Outlook OST file to Dovecot-compatible Maildir++ format.

    \b
    INPUT_OST       Path to the .ost (or .pst) file.
    OUTPUT_MAILDIR  Destination directory (created if absent). Becomes ~/Maildir/ on the server.
    """
    _setup_logging(verbose, log_file)

    click.echo(f"ost2maildir v1.0.0  |  {datetime.datetime.now():%Y-%m-%d %H:%M}")
    click.echo(f"Input : {input_ost}")
    click.echo(f"Output: {output_maildir}")

    reader = OSTReader(str(input_ost))

    # Encryption check
    try:
        enc = reader.detect_encryption()
        click.echo(f"Encryption: {enc}")
    except EncryptionError as exc:
        click.secho(f"\nERROR: {exc}", fg="red", err=True)
        sys.exit(1)

    # Dry run: count and exit
    if dry_run:
        click.echo("\nDry-run mode — counting items...")
        counts = reader.count_items()
        reader.close()
        click.echo(f"\n  Email messages : {counts['messages']:>8,}")
        click.echo(f"  Calendar items : {counts['calendar']:>8,}")
        click.echo(f"  Contacts       : {counts['contacts']:>8,}")
        click.echo(f"  Tasks/other    : {counts['other']:>8,}")
        click.echo(f"  Parse errors   : {counts['skipped_errors']:>8,}")
        click.echo(f"  TOTAL          : {sum(counts.values()):>8,}")
        est_gb = (counts['messages'] * 15_000 + counts['calendar'] * 2_000 + counts['contacts'] * 1_000) / 1e9
        click.echo(f"\n  Estimated output: ~{est_gb:.1f} GB")
        return

    output_maildir.mkdir(parents=True, exist_ok=True)

    writer = MaildirWriter(str(output_maildir))
    folder_mapper = FolderMapper()
    dedup = Deduplicator(str(output_maildir), enabled=not no_dedup)
    mime_converter = MIMEConverter(hostname=hostname, default_encoding=encoding)

    stats = {
        "written": 0,
        "duplicates": 0,
        "skipped_class": 0,
        "errors": 0,
        "calendar": 0,
        "contacts": 0,
    }

    click.echo("\nConverting...")

    with tqdm(unit="msg", dynamic_ncols=True, colour="green") as pbar:
        for folder_path, msg in reader.iter_items():
            pbar.set_description(
                f"{'/'.join(folder_path[-2:]) or 'root':30s}"
            )

            try:
                _process_item(
                    msg=msg,
                    folder_path=folder_path,
                    writer=writer,
                    folder_mapper=folder_mapper,
                    dedup=dedup,
                    mime_converter=mime_converter,
                    skip_contacts=skip_contacts,
                    skip_calendar=skip_calendar,
                    hostname=hostname,
                    stats=stats,
                )
            except Exception as exc:
                log.warning("Unhandled error processing item in %r: %s", folder_path, exc)
                stats["errors"] += 1

            pbar.update(1)

    reader.close()
    writer.finalize()
    dedup.save()

    # Final summary
    click.echo(f"\n{'─' * 50}")
    click.echo(f"  Messages written   : {stats['written']:>8,}")
    click.echo(f"  Calendar items     : {stats['calendar']:>8,}")
    click.echo(f"  Contacts           : {stats['contacts']:>8,}")
    click.echo(f"  Duplicates skipped : {stats['duplicates']:>8,}")
    click.echo(f"  Skipped (class)    : {stats['skipped_class']:>8,}")
    click.echo(f"  Errors             : {stats['errors']:>8,}")
    click.echo(f"{'─' * 50}")
    click.echo(f"\nOutput: {output_maildir}")
    click.echo("\nTo import into Dovecot:")
    click.echo(f"  doveadm import -u USERNAME maildir:{output_maildir.absolute()} \"\" \"*\"")


def _process_item(
    msg,
    folder_path: list,
    writer: MaildirWriter,
    folder_mapper: FolderMapper,
    dedup: Deduplicator,
    mime_converter: MIMEConverter,
    skip_contacts: bool,
    skip_calendar: bool,
    hostname: str | None,
    stats: dict,
):
    props = get_all_mapi_props(msg)
    mc = entry_as_string(props.get(PR_MESSAGE_CLASS)) or "IPM.Note"

    # --- Skip unwanted classes ---
    if mc.startswith("IPM.Task") or mc.startswith("IPM.StickyNote"):
        stats["skipped_class"] += 1
        return

    # --- Calendar ---
    if mc.startswith("IPM.Appointment") or mc.startswith("IPM.Schedule.Meeting"):
        if skip_calendar:
            stats["skipped_class"] += 1
            return
        mime_msg, flags = convert_calendar(msg, hostname)
        mid = mime_msg.get("Message-ID", "")
        if dedup.is_duplicate(mid):
            stats["duplicates"] += 1
            return
        ts = _message_timestamp(props)
        writer.write(".Calendar", mime_msg, flags, ts)
        stats["calendar"] += 1
        return

    # --- Contacts ---
    if mc.startswith("IPM.Contact") or mc.startswith("IPM.DistList"):
        if skip_contacts:
            stats["skipped_class"] += 1
            return
        mime_msg, flags = convert_contact(msg, hostname)
        mid = mime_msg.get("Message-ID", "")
        if dedup.is_duplicate(mid):
            stats["duplicates"] += 1
            return
        ts = _message_timestamp(props)
        writer.write(".Contacts", mime_msg, flags, ts)
        stats["contacts"] += 1
        return

    # --- Email (IPM.Note, Report.*, and anything else) ---
    # Skip pure internal/metadata classes
    if any(mc.startswith(p) for p in ("IPM.Configuration", "IPM.Note.Rules", "IPM.Note.StorageQuota")):
        stats["skipped_class"] += 1
        return

    mime_msg, flags = mime_converter.convert(msg)

    # Deduplication
    mid = mime_msg.get("Message-ID") or synthetic_message_id(msg, hostname)
    if not mime_msg.get("Message-ID"):
        mime_msg["Message-ID"] = mid
    if dedup.is_duplicate(mid):
        stats["duplicates"] += 1
        return

    maildir_name = folder_mapper.get_maildir_name(folder_path)
    ts = _message_timestamp(props)
    if writer.write(maildir_name, mime_msg, flags, ts):
        stats["written"] += 1
    else:
        stats["errors"] += 1


def _message_timestamp(props: dict) -> float:
    """Get Unix timestamp for Maildir filename from MAPI date properties."""
    for prop_id in (PR_MESSAGE_DELIVERY_TIME, PR_CLIENT_SUBMIT_TIME, PR_CREATION_TIME):
        entry = props.get(prop_id)
        if entry is None:
            continue
        try:
            dt = entry.data_as_datetime
            if dt:
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=datetime.timezone.utc)
                return dt.timestamp()
        except Exception:
            pass
    return datetime.datetime.now(datetime.timezone.utc).timestamp()


def _setup_logging(verbose: bool, log_file: Path | None):
    level = logging.DEBUG if verbose else logging.INFO
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stderr)]
    if log_file:
        handlers.append(logging.FileHandler(str(log_file)))
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=handlers,
    )
    # Quiet noisy libraries
    logging.getLogger("icalendar").setLevel(logging.WARNING)
