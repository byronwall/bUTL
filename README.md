#bUTL (Excel add-in)
bUTL is an add-in for Excel that started as an accumulation of utility and helper code collected by @byronwall over the years.  It is intended to continue in that vein, growing to include general utilities from others.

##Installation
To use this add-in, simply grab the current [bUTL.xlam](/bUTL.xlam) file from [Releases](/releases).

##Help
Documentation for the features of bUTL are located in the [/docs/ folder](/docs/README.md).  This contains a run down of all the features included in the Ribbon interface.  There are a number of `Subs` which are included in the add-in not placed on the Ribbon.  This is being resolved.

##Contributing
For purposes of development, it is assumed that the **current** source code for the add-in is contained in the [/src/](/src/) folder.  The compiled (really zipped) `bUTL.xlam` is no longer contained in the repo.

If you want to contribute a feature to the add-in or improve the code in some other way (e.g. fix a bug), please use the following workflow:

 - clone the repo
 - rebuild the xlam file from src, see `scripts/create xlam from src`
 - make changes using the VBA editor and possibly the Ribbon editor, saving the file like normal
 - export your new add-in back to src, see `scripts/create src from xlam`
 - verify that the diffs on the files inside src seem reasonable, there might be a number of xml files and vba/frx files that are generated which were not actually changed; please don't commit these
 - commit and submit a pull request

Why? This workflow has been adopted because Excel/VBA files have severe limitations for version control.  An `xlam` file is a zipped folder with a number of binary files inside the zip.  There are a couple of useful add-ins which help manage this, but I would prefer to not dictate what add-ins are installed.  Given that, this workflow and build scripts allow for changes to the underlying VBA code and Ribbon interface to be properly tracked.
