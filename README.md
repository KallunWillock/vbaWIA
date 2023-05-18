# vbaWIA
A collection of routines and articles I've written and/or collated about the Windows Imaging Acquisition COM Object


## Image Properties

### :framed_picture: Basic

WIA provides an easy way to get access to basic images properties on the one hand, and more to both read and write more advanced image properties, such as EXIF MetaData and Animated GIF frames. Below, I set out a simple routine to demonstrate how to load an image into the WIA Object, and then access the various properties. This is not intended to be exhaustive. My suggestion would be to step through the code, and check out the WIA Object in the Object Browser in the VBA IDE once the image has loaded in order to get a good understanding.

The routines set out in Part 1 demonstrate how to convert various data sources into an stdPicture object. In order to get access to the image properties, though, it is necessary to convert the data into an ImageFile - one of the three main WIA Objects used for image manipulation/display in VBA. (the other two are Vector and ImageProcess). As such, I've had to write an additional function - GetImageFileFromURL (which bears a stunning similarity to one of the Part 1 functions - GetImageFromURL); this will read an image file from a URL.

Here is a test sub with that will process two files: (1) an image on the internet (I have used this one before, because it has EXIF Metadata on it); and
(2) an animated GIF file of a dancing panda... because why not. The static image is on the right hand side for your reference. The code envisages that you download the GIF image yourself (URL in the comments) so as to demonstrate how to load an image from your local drive using the LoadFile method, but you're welcome to use another image.

### :world_map: EXIF MetaData
https://www.mrexcel.com/board/threads/using-a-userform-to-change-the-document-properties-or-tags.1198206/

First, it would be helpful if we can check to make sure that this approach will work, so can you please try the Test_ReadProperties code with this sample photo. You will need to change the path to the file below. I have added EXIF metadata to image, the code I used for doing that is set out below in one of the test subroutines.

Please note that the WriteEXIFData will, by default, write over the image file you give it. I have built in an option for it not to do that, in which case it will return the new filename as a string. Also, it includes an option to add a backup. As you can probably tell, I'm super sensitive about ever deleting files with VBA. That being the case, please can you be sure to test this on dummy image files (like the one above) and make sure that you have backups of everything. The default assumption is the there is no undo available.

The functionality is a bit limited, but I just wanted to check to make sure it'll work. Do let me know if you're having any troubles with it. I have family visiting this weekend, but I should generally be around if you need help. Fingers crossed!
![Sample exif](https://github.com/KallunWillock/JustMoreVBA/blob/main/Images/pexels-jill-evans-11567527.jpg)

### Transforms
https://www.mrexcel.com/board/threads/resize-image-inside-a-image-control-on-userform.1234401/
The WIA COM object can scale images, rotate them, and save them as different image formats. 

The following class module does those three things - note that this is a very basic example, and doesn't include full error handling, etc. Basically, you need to load a picture by setting the the class Source property to the filename. 

clsWIA comprises the following methods:

* :arrows_counterclockwise: RotateImage - rotate the image to either 0, 90, 180, or 270 degrees. 
* :straight_ruler: ScaleImage - there is a further MaintainAspect parameter (defaults to True) which maintains the aspect ratio of the original image, but you can set it to False if you really, really, really want to change the image to 100x100. 
* SaveImage - save/convert with a file format parameter set to PNG file format.

![Sample image](https://github.com/KallunWillock/vbaWIA/blob/main/Transforms/WIA_RotateScale.png)