# import packages
import os, shutil, logging

logging.basicConfig(
    level = logging.INFO,
    format = '[%(levelname)s] %(asctime)s — %(message)s',
    datefmt = '%Y-%m-%d %H:%M:%S'
)


# set the file path
src_folder = 'utils'
dst_folder = 'utils_local'

# make the folder
os.makedirs(dst_folder, exist_ok=True)

# copy the .py files to the utils_local
for file_name in os.listdir(src_folder):
    src_path = os.path.join(src_folder, file_name)
    dst_path = os.path.join(dst_folder, file_name)

    if os.path.isfile(src_path) and file_name.endswith('.py'):
        shutil.copy2(src_path, dst_path)

logging.info("utils_local generated")
