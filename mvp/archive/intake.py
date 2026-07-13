import os
import gdown


def handle_intake(excel_file, zip_file, drive_link, upload_folder):
    os.makedirs(upload_folder, exist_ok=True)

    excel_path = os.path.join(upload_folder, 'candidates.xlsx')
    excel_file.save(excel_path)

    zip_path = os.path.join(upload_folder, 'input.zip')

    if zip_file:
        zip_file.save(zip_path)
    elif drive_link:
        try:
            gdown.download(drive_link, zip_path, quiet=False, fuzzy=True)
        except Exception as e:
            raise Exception(
                f'Google Drive download failed: {e}. '
                'Make sure the link is set to "Anyone with the link can view".'
            )
        if not os.path.exists(zip_path) or os.path.getsize(zip_path) == 0:
            raise Exception('Downloaded file is empty. Check Drive link permissions.')
    else:
        raise Exception('No ZIP file or Drive link provided.')

    return zip_path, excel_path
