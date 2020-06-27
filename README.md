# Outlook-Late-Binding
How to use the outlook COM library reference from Excel without early binding

When automating Outlook from Excel you should use late binding rather than earlier binding. Outlook changes the library name every time there is a version upgrade. So every time a user is upgraded each macro they use to automate Outlook has to be modified. Sime some offices have people on multiple versions of Outlook that use the same file, using early binding will force you to choose which user it will work for. 
Late binding is explained in the comments. 

Also, there are two versions of the macro. One as text as one as an image. From your email it sounds like you'll need to use the image version. That's a little complicated because the Outlook object model doesn't allow you to paste an Excel range into an Outlook email using VBA. If you write the call in Outlook there may be a way to use OLE to embed an excel range in an email, but I think I looked for one and couldn't find it. It's quite easy to do by hand. 
