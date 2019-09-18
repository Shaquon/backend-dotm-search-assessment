#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Shaquon"

import os
import sys
import argparse
import zipfile

#takes 
def create_parser():
    parser = argparse.ArgumentParser(description='searches through dotm files for text')
    parser.add_argument('--dir', help='the directory you want to search through', default='.')
    parser.add_argument('text', help='this is the text that you want to search for')
    return parser
# args = parser.parse_args()
# print(args.accumulate(args.integers))

def search_file_for_text(filename, searchtext):
    """1) opens up a if file is a zipfile if not error is printed
       2) searches for substring"""
    if zipfile.is_zipfile(filename) == False:
        print("error: not a zipfile")
        return 0
    with zipfile.ZipFile(filename) as zip:
        # print(zip.namelist())
        with zip.open('word/document.xml') as doc:
                for line in doc:
                    indexOfSearchText = line.find(searchtext)
                    if indexOfSearchText >= 0:
                        print("Match found in file: {}".format(filename))
                        print("Text Context: " + line[indexOfSearchText-40:indexOfSearchText+40]+" ")
                        return True
                return False


def main():
   # raise NotImplementedError("Your awesome code begins here!")
    """ in this function we want to use the parse and search_file_for_text functions from above. We will first start off by """
    parser = create_parser()
    files_searched = 0
    matches_found = 0
    program_args = parser.parse_args()
    file_directory = os.listdir(program_args.dir)
    print("search through directory for text")
    for document in file_directory:
        files_searched += 1
        path = os.path.join(program_args.dir, document)
        if not document.endswith(".dotm"):
            print("skipped")
            continue
        if search_file_for_text(path, program_args.text):
            matches_found += 1
    print('files searched through ={}'.format(files_searched))
    print('files matched ={}'.format(matches_found))


if __name__ == '__main__':
    main()
