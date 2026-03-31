import os
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS
import datetime
import shutil
import time

samsung = ('./photos/samsung', 'SamsungS23')
iphone = ('./photos/iphone', 'iPhone')

excel_path = 'rename_test.xlsx'
new_excel_path = 'new_db.xlsx'
data_sheet_name = 'Data'


def exif_extraction(file_path) -> list:
  exif_dict = []
  exif_dict.append(os.path.basename(file_path))
  # print(exif_dict)
  try:
    pil_img = Image.open(file_path)
    exif_info = pil_img._getexif()
    if exif_info:
      exif = {TAGS.get(k, k): v for k, v in exif_info.items()}

      if "DateTimeOriginal" in exif:
        dt_obj = datetime.datetime.strptime(exif["DateTimeOriginal"], '%Y:%m:%d %H:%M:%S')
        exif_dict.append(dt_obj.timestamp())

      elif "DateTime" in exif:
        dt_obj = datetime.datetime.strptime(exif["DateTime"], '%Y:%m:%d %H:%M:%S')
        exif_dict.append(dt_obj.timestamp())

      if "FocalLength" in exif:
        exif_dict.append(exif["FocalLength"])

  except Exception as e:
    print(f"Warning: Could not read EXIF data for {file_path}. Error: {e}")
    pass

  return exif_dict

def bulk_rename(path_options, excel_path, sheet_name='Data') -> list:
    exif_details_map = {}
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    valid_rows = df.loc[df['Modello_Telefono'] == path_options[1]]

    target_names = valid_rows['ID_Foto'].dropna().tolist()
    print(f"ID_Foto from Excel: {target_names}")

    valid_extensions = ('.jpg', '.jpeg', '.png', '.JPG')
    all_files_in_dir = [f for f in os.listdir(path_options[0]) if f.endswith(valid_extensions) and not f.startswith('.')]

    files_with_exif = []
    files_without_exif = []

    print("Processing files to check EXIF data...")
    for filename in all_files_in_dir:
        file_path = os.path.join(path_options[0], filename)
        exif_result = exif_extraction(file_path)
        original_filename = exif_result[0]

        acquisition_time = None
        focal_length = None

        if len(exif_result) > 1:
            acquisition_time = exif_result[1]
        if len(exif_result) > 2:
            focal_length = exif_result[2]

        exif_details_map[original_filename] = {'timestamp': acquisition_time, 'focal_length': focal_length}

        if acquisition_time is not None:
            files_with_exif.append((filename, acquisition_time))
        else:
            files_without_exif.append(filename)

    # Sort files that have EXIF data by their acquisition time
    files_with_exif.sort(key=lambda x: x[1])
    sorted_filenames = [f[0] for f in files_with_exif]

    if files_without_exif:
        print(f"The following {len(files_without_exif)} files could not be processed due to missing or problematic EXIF data and require manual handling:")
        for f in files_without_exif:
            print(f" - {f}")

    if len(sorted_filenames) != len(target_names):
        print(f"ERROR: ({len(sorted_filenames)}) files with valid EXIF != ({len(target_names)}) excel ids")
        return []

    confirm = input("Proceed with renaming files with valid EXIF data? (y/n) ")
    if confirm.lower() != 'y':
        return []

    os.makedirs(os.path.join(path_options[0], "renamed"), exist_ok=True)

    excel_update_data = []

    for i, filename in enumerate(sorted_filenames):
        original_filename = filename
        extension = os.path.splitext(filename)[1].lower()
        new_name_base = target_names[i]
        new_filename = os.path.splitext(new_name_base)[0] + extension

        exif_data_for_original = exif_details_map.get(original_filename, {})
        focal_length_for_excel = exif_data_for_original.get('focal_length')

        excel_update_data.append({
            'ID_Foto': new_filename,
            'Focale_EXIF': str(focal_length_for_excel) if focal_length_for_excel is not None else None
        })

        old_path = os.path.join(path_options[0], filename)
        new_path = os.path.join(path_options[0], "renamed", new_filename)

        try:
            shutil.copy2(old_path, new_path)
            # print(f"Copiato e rinominato: {filename} -> {new_filename}")
        except Exception as e:
            print(f"Errore su {filename}: {e}")
    print(excel_update_data)
    return excel_update_data

def update_excel_with_exif_data(excel_path, output_excel_path, sheet_name, exif_data) -> None:
    print(exif_data)
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"Error: Excel file not found at {excel_path}")
        return

    df['Focale_EXIF'] = df['Focale_EXIF'].astype('object')

    if not output_excel_path:
      output_excel_path = excel_path

    for entry in exif_data:
        new_filename = entry.get('ID_Foto')
        focal_length = entry.get('Focale_EXIF')

        if new_filename and focal_length is not None:
            mask = df['ID_Foto'] == new_filename
            if mask.any():
                df.loc[mask, 'Focale_EXIF'] = focal_length
            else:
                print(f"Warning: ID_Foto '{new_filename}' not found in Excel for EXIF update.")

    try:
        with pd.ExcelWriter(output_excel_path, engine='xlsxwriter', mode='w') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Excel file '{output_excel_path}' updated successfully.")
    except Exception as e:
        print(f"Error saving updated Excel file: {e}")





start_time = time.time()
rows = bulk_rename(samsung, excel_path, data_sheet_name)
if rows:
    update_excel_with_exif_data(excel_path, new_excel_path, data_sheet_name, rows)
else:
    print("No data to update Excel file.")
print("--- %s seconds ---" % (time.time() - start_time))