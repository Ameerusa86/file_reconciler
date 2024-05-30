import os
import pandas as pd
from datetime import datetime
import smtplib
from email.message import EmailMessage


def list_files(directory):
    """List files in the given directory."""
    try:
        return sorted(os.listdir(directory))
    except FileNotFoundError:
        print(f"Directory {directory} does not exist.")
        return []


def compare_directories(dir1, dir2):
    """Compare files in two directories and return the differences."""
    files_dir1 = set(list_files(dir1))
    files_dir2 = set(list_files(dir2))

    only_in_dir1 = files_dir1 - files_dir2
    only_in_dir2 = files_dir2 - files_dir1

    return (
        len(files_dir1),
        len(files_dir2),
        len(only_in_dir1),
        len(only_in_dir2),
        only_in_dir1,
        only_in_dir2,
    )


def generate_excel_report(
    dir1, dir2, num_files_dir1, num_files_dir2, only_in_dir1, only_in_dir2, report_file
):
    """Generate a reconciliation report in an Excel file."""
    report_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Create a DataFrame for the report summary
    summary_data = {
        "Folder": ["ESO", "EMSmart"],
        "Number of Files": [num_files_dir1, num_files_dir2],
    }
    summary_df = pd.DataFrame(summary_data)

    # Create a DataFrame for missing files in EMSmart folder
    missing_emsmart_df = pd.DataFrame(
        only_in_dir1, columns=["Missing Files in EMSmart"]
    )

    with pd.ExcelWriter(report_file, engine="xlsxwriter") as writer:
        # Write report summary
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # Write report date
        summary_sheet = writer.sheets["Summary"]
        summary_sheet.write(0, 4, "Report Date")
        summary_sheet.write(1, 4, report_date)

        # Write missing files in EMSmart folder
        missing_emsmart_df.to_excel(
            writer, sheet_name="Missing Files in EMSmart", index=False
        )

        # Add formatting
        workbook = writer.book
        header_format = workbook.add_format(
            {"bold": True, "font_color": "white", "bg_color": "#2E75B6"}
        )
        summary_sheet.set_row(0, None, header_format)
        summary_sheet.set_column(0, 1, 20)

    print(f"Reconciliation report generated: {report_file}")
    return report_date


# def send_email_with_attachment(
#     report_date,
#     recipient_email,
#     report_file,
#     sender_email,
#     sender_password,
#     smtp_server,
#     smtp_port,
# ):
#     """Send an email with the reconciliation report attached."""
#     msg = EmailMessage()
#     msg["Subject"] = f"Reconciliation Report - {report_date}"
#     msg["From"] = sender_email
#     msg["To"] = recipient_email
#     msg.set_content(
#         f"Please find the reconciliation report attached.\nReport Date: {report_date}"
#     )

#     with open(report_file, "rb") as f:
#         file_data = f.read()
#         file_name = os.path.basename(report_file)

#     msg.add_attachment(
#         file_data, maintype="application", subtype="octet-stream", filename=file_name
#     )

#     with smtplib.SMTP_SSL(smtp_server, smtp_port) as smtp:
#         smtp.login(sender_email, sender_password)
#         smtp.send_message(msg)


if __name__ == "__main__":
    directory1 = r"K:\Python\RPA\File_Reconciliation\ESO"  # ESO directory
    directory2 = r"K:\Python\RPA\File_Reconciliation\EMSmart"  # EMSmart directory
    report_file = "reconciliation_report.xlsx"
    recipient_email = "ameer.hasan@emsmc.com"
    sender_email = "engamermecha@gmail.com"  # Update with your Gmail address
    sender_password = "$@Saralydia1986$#"  # Update with your Gmail password
    smtp_server = "smtp.gmail.com"
    smtp_port = 465

    num_files_dir1, num_files_dir2, num_missing_emsmart, _, only_in_dir1, _ = (
        compare_directories(directory1, directory2)
    )

    report_date = generate_excel_report(
        directory1,
        directory2,
        num_files_dir1,
        num_files_dir2,
        only_in_dir1,
        [],
        report_file,
    )

    # send_email_with_attachment(
    #     report_date,
    #     recipient_email,
    #     report_file,
    #     sender_email,
    #     sender_password,
    #     smtp_server,
    #     smtp_port,
    # )
