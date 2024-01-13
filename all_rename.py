import os
import datetime
import pytz
import random
import string

from PIL import Image
from win32com.propsys import propsys, pscon

def get_file_dates(path):
    dates = []

    # Check if "Date Taken" field is available
    try:
        with Image.open(path) as img:
            if hasattr(img, '_getexif'):
                exif_data = img._getexif()
                if exif_data is not None and 36867 in exif_data:
                    dates.append(datetime.datetime.strptime(exif_data[36867], "%Y:%m:%d %H:%M:%S"))
    except:
        pass

    # Get file creation time
    creation_time = datetime.datetime.fromtimestamp(os.path.getctime(path))
    dates.append(creation_time)

    # Get file modification time
    modification_time = datetime.datetime.fromtimestamp(os.path.getmtime(path))
    dates.append(modification_time)

    # Return the oldest date among all the available options
    return min(dates)

def get_media_created_time(path):
    dt = datetime.datetime.now()

    try:
        properties = propsys.SHGetPropertyStoreFromParsingName(path)
        dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()
        if not isinstance(dt, datetime.datetime):
            # In Python 2, PyWin32 returns a custom time type instead of
            # using a datetime subclass. It has a Format method for strftime
            # style formatting, but let's just convert it to datetime:
            dt = datetime.datetime.fromtimestamp(int(dt))
    except:
        pass

    return dt

def rename_video_files(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            if file_path.lower().endswith(('.mov', '.mp4')):
                media_created = get_media_created_time(file_path)
                if media_created is None:
                    media_created = datetime.datetime.now()
                unique_identifier = generate_unique_identifier()
                new_name = media_created.strftime("%Y-%m-%d_%H%M%S") + "_" + unique_identifier + os.path.splitext(filename)[1]
                new_path = os.path.join(directory, new_name)
                os.rename(file_path, new_path)
                print(f"Renamed: {filename} to {new_name}")
            else:
                date = get_file_dates(file_path)
                if date is None:
                    date = datetime.datetime.now()
                unique_identifier = generate_unique_identifier()
                new_filename = date.strftime("%Y-%m-%d_%H%M%S") + "_" + unique_identifier + os.path.splitext(filename)[1]
                new_file_path = os.path.join(directory, new_filename)
                os.rename(file_path, new_file_path)
                print(f"Renamed {filename} to {new_filename}")

def rename_image_files(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            if file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
                date = get_file_dates(file_path)
                if date is None:
                    date = datetime.datetime.now()
                unique_identifier = generate_unique_identifier()
                new_filename = date.strftime("%Y-%m-%d_%H%M%S") + "_" + unique_identifier + os.path.splitext(filename)[1]
                new_file_path = os.path.join(directory, new_filename)
                os.rename(file_path, new_file_path)
                print(f"Renamed {filename} to {new_filename}")

                if file_path.lower().endswith(('.jpeg','.jpg')):
                    video_filename = os.path.splitext(filename)[0] + '.mov'
                    new_video_filename = os.path.splitext(new_filename)[0] + '.mov'
                    old_file_path = os.path.join(directory, video_filename)
                    new_file_path = os.path.join(directory, new_video_filename)
                    try:
                        os.rename(old_file_path, new_file_path)
                        print(f"Renamed {video_filename} to {new_video_filename}")
                    except:
                        pass

def rename_special_files(directory, date):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            unique_identifier = generate_unique_identifier()
            new_filename = date + "_" + unique_identifier + os.path.splitext(filename)[1]
            new_file_path = os.path.join(directory, new_filename)
            os.rename(file_path, new_file_path)
            print(f"Renamed {filename} to {new_filename}")


def generate_unique_identifier(k=4):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=k))
    #return str(random.randint(1000, 9999))

# Example usage
#directory_path = "F:\\1 - Photos\\2021\\11\\mp4"
#directory_path = "F:\\1 - Photos\\2023-03\\iCloud Photos"
directory_path = "C:\\Users\\Admin\\Desktop\\Photos and Videos\\iCloud Photos"

#rename_image_files(directory_path) # standard, step 1
#rename_video_files(directory_path) # movies, step 2
rename_special_files(directory_path, "2023_07_00") # non-conforming, step 3
