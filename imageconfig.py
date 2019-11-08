import configparser
import os
from collections import namedtuple


config = configparser.ConfigParser()
config_file = 'image_config.ini'
config.read(config_file)


def has_parent(path):

    par_dir = os.path.abspath(os.path.join(path, os.pardir))
    return par_dir, os.path.exists(par_dir)


def mkdir_if_parent_present(path):

    par_dir, present = has_parent(path)
    if present:
        if not os.path.exists(path):
            os.mkdir(path)
        return path
    else:
        raise FileNotFoundError(f'Please ensure that {par_dir} exists. Check {config_file}')


# image to excel


INPUT_IMAGE_DIR = mkdir_if_parent_present(config['IMAGE']['INPUT_IMAGE_DIR'])
OUTPUT_IMG_ZIP_NAME = config['IMAGE']['OUTPUT_IMG_ZIP_NAME']
OUTPUT_IMG_EXCEL_DIR = mkdir_if_parent_present(
    config['IMAGE']['OUTPUT_IMG_EXCEL_DIR'])
OUTPUT_IMG_ZIP = mkdir_if_parent_present(
    config['IMAGE']['OUTPUT_IMG_ZIP'])
OUTPUT_LOG_DIR = mkdir_if_parent_present(config['IMAGE']['OUTPUT_LOG_DIR'])
OUTPUT_ERROR_DIR = mkdir_if_parent_present(config['IMAGE']['OUTPUT_ERROR_DIR'])

image_tuple = namedtuple('IMAGE', [
                         'INPUT_IMAGE_DIR', 'OUTPUT_IMG_ZIP_NAME', 'OUTPUT_IMG_EXCEL_DIR', 'OUTPUT_IMG_ZIP', 'OUTPUT_LOG_DIR','OUTPUT_ERROR_DIR'])
IMAGE = image_tuple(INPUT_IMAGE_DIR, OUTPUT_IMG_ZIP_NAME,
                    OUTPUT_IMG_EXCEL_DIR, OUTPUT_IMG_ZIP,OUTPUT_LOG_DIR,OUTPUT_ERROR_DIR)


# PROCESSING_OUTPUT = config['DEFAULT']['PROCESSING_OUTPUT']

__all__ = [IMAGE]

if __name__ == "__main__":

    print(IMAGE)