The WIA COM object can scale images, rotate them, and save them as different image formats. 

The following class module does those three things - note that this is a very basic example, and doesn't include full error handling, etc. Basically, you need to load a picture by setting the the class Source property to the filename. 

You can use RotateImage routine to rotate the image to either 0, 90, 180, or 270 degrees. 
You can ScaleImage to 100x100 (as per your instructions above) - there is a further MaintainAspect parameter (defaults to True) which maintains the aspect ratio of the original image, but you can set it to False if you really, really, really want to change the image to 100x100. 
The demo routine below has invokes the SaveImage routine with a file format parameter set to PNG file format.

![Sample image](KallunWillock/vbaWIA/Transforms/WIA_RotateScale.png)