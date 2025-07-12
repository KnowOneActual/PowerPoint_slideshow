Automated PowerPoint Slideshow GeneratorA Python script to instantly turn a folder of images into a polished, widescreen PowerPoint presentation. Perfect for event technicians, marketing teams, or anyone who needs to quickly create a looping slideshow from a batch of photos.The Problem It SolvesEver been handed a flash drive with hundreds of photos moments before an event, with the request to "get these on the big screen"? This script automates that entire process, saving you from the tedious manual work of inserting, resizing, and centering each image in PowerPoint.Key Features✅ Smart Image Processing: Automatically verifies that files are valid images, skipping corrupt files or non-image documents.🔄 Auto-Rotation: Reads camera EXIF data to correctly orient photos, so portrait shots don't end up sideways.⚡️ Optimized for Screens: Resizes images to a maximum of 1920x1080 (Full HD), drastically reducing the final .pptx file size without sacrificing quality on screen.🎨 Standardized Format: Converts all images (including PNGs with transparency) to a consistent JPEG format.✨ Polished Output: Creates a clean, 16:9 widescreen presentation with each image perfectly centered.🧹 Automatic Cleanup: Uses a temporary folder for processing and deletes it when finished, keeping your original folder clean.RequirementsYou'll need to have Python 3 installed on your system. You can install the necessary libraries using pip:pip install Pillow python-pptx
How to UseDownload the Script: Save the create_slideshow.py file to your computer.Run it from the Terminal: Open a terminal or command prompt, navigate to the directory where you saved the script, and run it with the following command:python create_slideshow.py
Provide the Folder Path: The script will prompt you to enter the full path to the folder containing your images.Enter the full path to the folder containing your images: C:\Users\YourUser\Desktop\EventPhotos
Done!: The script will process the images and save a new file named Generated_Slideshow.pptx to your Desktop.Example OutputProcessing images from: C:\Users\YourUser\Desktop\EventPhotos
Temporary files will be stored in: C:\Users\YourUser\Desktop\EventPhotos\processed_images

Found 150 images. Starting processing...
  ✅ Processed: IMG_001.jpg
  ✅ Processed: IMG_002.png
  ⚠️ Skipping 'notes.docx': Not a valid or supported image file.
  ✅ Processed: IMG_003.jpeg
  ...

All images processed. Creating PowerPoint presentation...

✨ Success! Slideshow saved to your Desktop as 'Generated_Slideshow.pptx'

🧹 Cleaned up temporary files.
LicenseThis project is licensed under the MIT License. See the LICENSE file for details.
