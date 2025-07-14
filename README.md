### SlideWrangler: Automated PowerPoint Slideshow Generator

A Python script to instantly turn a folder of images into a polished, widescreen PowerPoint presentation.

This tool is featured in my blog post:[From Flash Drive Frenzy to Slideshow Serenity](https://blog.beaubremer.com/posts/slideshow_serenity/)



## The Problem It Solves

Ever been handed a flash drive with hundreds of photos moments before an event, with the request to "get these on the big screen"? This script automates that entire process, saving you from the tedious manual work of inserting, resizing, and centering each image in PowerPoint.


## Key Features



* ** Smart Image Processing**: Automatically verifies that files are valid images, skipping corrupt files or non-image documents.
* ** Auto-Rotation**: Reads camera EXIF data to correctly orient photos, so portrait shots don't end up sideways.
* ** Optimized for Screens**: Resizes images to a maximum of 1920x1080 (Full HD), drastically reducing the final .pptx file size without sacrificing quality on screen.
* ** Standardized** Format**: Converts all images (including PNGs with transparency) to a consistent JPEG format.
* ** Polished Output**: Creates a clean, 16:9 widescreen presentation with each image perfectly centered.
* ** Automatic Cleanup**: Uses a temporary folder for processing and deletes it when finished, keeping your original folder clean.


## Requirements

You'll need to have Python 3 installed on your system. You can install the necessary libraries using pip:

pip install Pillow python-pptx \



## How to Use



1. **Download the Script**: Save the create_slideshow.py file to your computer.
2. **Run it from the Terminal**: Open a terminal or command prompt, navigate to the directory where you saved the script, and run it with the following command: \
python create_slideshow.py \

3. **Provide the Folder Path**: The script will prompt you to enter the full path to the folder containing your images. \
Enter the full path to the folder containing your images: C:\Users\YourUser\Desktop\EventPhotos \

4. **Done!**: The script will process the images and save a new file named Generated_Slideshow.pptx to your Desktop.

