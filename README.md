# python_image_sort
At my current job, one typically has to process sometimes hundreds of images in Excel. This involves moving the
image to the Excel sheet, shrinking it, aligning it in a cell, then copying the name of the image over and 
ensuring it's correct. This typically takes over an hour to do, and the task is performed almost daily. 
I wrote a simple python script which automated this task. This script will pull all files and names from a chosen
directory, sort the images/filenames by alphabetical order, and insert them into a new Excel document with whatever 
size image scaling you choose.

**Note, several variables and pathnames must be updated for the script to work. The names of these locations have been
marked at the start of the document as well as surrounded by comments in the code.
