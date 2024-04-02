# utl-perl-write-and-read-meta-data-saved-in-the-windows-file-properties-details-panel
    %let pgm=utl-perl-write-and-read-meta-data-saved-in-the-windows-file-properties-details-panel;

    Write and read meta data saved in the windows file properties details panel

    Two Solutions
        1 perl pop get meta data (could not do it in R, Python or Powershell)
        2 posershell get meta (does not seem to relate to the details panel)

    github
    https://tinyurl.com/mrx2hzsr
    https://github.com/rogerjdeangelis/utl-perl-write-and-read-meta-data-saved-in-the-windows-file-properties-details-panel

    inspired by
    https://goo.gl/UbGFGd
    https://communities.sas.com/t5/Base-SAS-Programming/author-of-excel-file-importing-to-SAS/m-p/338811


    BUG:
      Although SAS can populate Property Categories perl Win3::OLE cannot read it back.

    Also
      Author and Name do not appear in the details panel but are populated and can be read.


    I prefer to use perl for batch scripting and windows/unices internals;

    PERL REPO
    -------------------------------------------------------------------------------------------------------------------------------
    https://github.com/rogerjdeangelis/utl-cross-patform-perl-script-to-parse-a-config-file-and-pass-parameters-to-SAS-batch
    https://github.com/rogerjdeangelis/utl-examples-of-drop-downs-from-sas-to-wps-r-microsoftR-python-perl-powershell
    https://github.com/rogerjdeangelis/utl-leveraging-your-knowledge-of-perl-regex-to-sas-wps-r-python-and-perl
    https://github.com/rogerjdeangelis/utl_pass_data_to_from_perl_R_python

    FYI you can create custom properties.

    /*                   _         _   _     _        __                        __ _ _            _      _        _ _
      ___ _ __ ___  __ _| |_ ___  | |_| |__ (_)___   / _|_ __ ___  _ __ ___    / _(_) | ___    __| | ___| |_ __ _(_) |___
     / __| `__/ _ \/ _` | __/ _ \ | __| `_ \| / __| | |_| `__/ _ \| `_ ` _ \  | |_| | |/ _ \  / _` |/ _ \ __/ _` | | / __|
    | (__| | |  __/ (_| | ||  __/ | |_| | | | \__ \ |  _| | | (_) | | | | | | |  _| | |  __/ | (_| |  __/ || (_| | | \__ \
     \___|_|  \___|\__,_|\__\___|  \__|_| |_|_|___/ |_| |_|  \___/|_| |_| |_| |_| |_|_|\___|  \__,_|\___|\__\__,_|_|_|___/

    */

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* DSD DATASET WANT total obs=8                                                                                           */
    /*                                                                                                                        */
    /* Obs    PROPERTY    VALUE                                                                                               */
    /*                                                                                                                        */
    /*  1     File        d\xls\roger.xlsx                                                                                    */
    /*  2     Author      Roger DeAngelis                                                                                     */
    /*  3     Title       Demographics for Target Protocol                                                                    */
    /*  4     Subject     Pubs                                                                                                */
    /*  5     Tags        SDTM clinical                                                                                       */
    /*  6     Categories                     bug in WIN32 Internals?  cannot read this property from file detail panel        */
    /*  7     Comments    Rogers Comments                                                                                     */
    /*  8     Name        roger.xlsx                                                                                          */
    /*                                                                                                                        */
    /**************************************************************************************************************************/


    /*                   _                                  _                    _              _       _
    / |  _ __   ___ _ __| |  _ __   ___  _ __     __ _  ___| |_   _ __ ___   ___| |_ __ _    __| | __ _| |_ __ _
    | | | `_ \ / _ \ `__| | | `_ \ / _ \| `_ \   / _` |/ _ \ __| | `_ ` _ \ / _ \ __/ _` |  / _` |/ _` | __/ _` |
    | | | |_) |  __/ |  | | | |_) | (_) | |_) | | (_| |  __/ |_  | | | | | |  __/ || (_| | | (_| | (_| | || (_| |
    |_| | .__/ \___|_|  |_| | .__/ \___/| .__/   \__, |\___|\__| |_| |_| |_|\___|\__\__,_|  \__,_|\__,_|\__\__,_|
     _  |_|              _  |_|         |_|      |___/
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

     %utlfkil(d:/xls/roger.xlsx);
     ods excel file="d:/xls/roger.xlsx"
     Title="Demographics for Target Protocol"
     Subject="Pubs"
     Keywords="SDTM clinical"
     Category="Demographics"
     Comments = "Rogers Comments"
     Author="Roger DeAngelis";
     proc print data=sashelp.class
       (obs=2 keep=name sex age);
     run;quit;
     ods excel close;

     d:/xls/roger.xlsx

      +----------------------+
      |     A   |  B   |  C  |
      +----------------------+
    1 |  NAME   | SEX  | AGE |
      +---------+------+-----+
    2 | ALFRED  |  M   | 14  |
      +---------+------+-----+
    3 | ALICE   |  F   | 15  |
      +---------+------+-----+

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    Right click on file select properties > details

     _________________________________________________________
    |                                                        |
    |                                                        |
    |  Right click on file select properties > details       |
    |                                                        |
    |  Roger.xlsx Properties                                 |
    |                                                        |
    |   +-------------------------------------------------+  |
    |   | General| Security | Details | Previous Versions |  |
    |   +-------------------------------------------------+  |
    |                                                        |
    |     Property    Value                                  |
    |                                                        |
    |     Discription ---------------------------------------|
    |                                                        |
    |     Title:      Demographics for Target Protocol       |
    |     Subject:    Pubs                                   |
    |     Tags:       SDTM clinical                          |
    |     Categories  Demographics                           |
    |     Comments:   Rogers Comments                        |
    |     Author:     Roger Deangelis                        |
    |                                                        |
    |________________________________________________________|

    /*               _                              _       _                 _
     _ __ ___   __ _| | _____   ___  __ _ ___    __| | __ _| |_ __ _ ___  ___| |_
    | `_ ` _ \ / _` | |/ / _ \ / __|/ _` / __|  / _` |/ _` | __/ _` / __|/ _ \ __|
    | | | | | | (_| |   <  __/ \__ \ (_| \__ \ | (_| | (_| | || (_| \__ \  __/ |_
    |_| |_| |_|\__,_|_|\_\___| |___/\__,_|___/  \__,_|\__,_|\__\__,_|___/\___|\__|

    */

    /*----  backticks are used to parse per lines                            ----*/

     %utl_submit_pl64('
     use strict;`
     use Win32::OLE;`
     my $file = "d:\\xls\\roger.xlsx";
     my $obj = Win32::OLE->GetObject($file);`
     my $title = $obj->{"Title"};`
     my $subject = $obj->{"Subject"};`
     my $tags = $obj->{"Keywords"};`
     my $category = $obj->{"Categories"};`
     my $comments = $obj->{"Comments"};`
     my $author = $obj->{"Author"};`
     my $name = $obj->{"Name"};`

     open (FH, ">", "d:/txt/roger.txt");`
     print FH "File: $file\n";`
     print FH "Author:$author\n";`
     print FH "Title: $title\n";`
     print FH "Subject: $subject\n";`
     print FH "Tags:$tags\n";`
     print FH "Category:$category\n";`
     print FH "Comments: $comments\n";`
     print FH "Name: $name\n";`
     ');

     data want;
      infile "d:/txt/roger.txt";
      input;
      put _infile_;
      Property = left(scan(_infile_,1,':'));
      Value    = left(cats(scan(_infile_,2,':'),scan(_infile_,3,':')));
     run;quit;

    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* WANT total obs=8                                                                                                       */
    /*                                                                                                                        */
    /* Obs    PROPERTY    VALUE                                                                                               */
    /*                                                                                                                        */
    /*  1     File        d\xls\roger.xlsx                                                                                    */
    /*  2     Author      Roger DeAngelis                                                                                     */
    /*  3     Title       Demographics for Target Protocol                                                                    */
    /*  4     Subject     Pubs                                                                                                */
    /*  5     Tags        SDTM clinical                                                                                       */
    /*  6     Category                    /* bug in WIN32 Internals?    */                                                    */
    /*  7     Comments    Rogers Comments                                                                                     */
    /*  8     Name        roger.xlsx                                                                                          */
    /*                                                                                                                        */
    /**************************************************************************************************************************/


    /*___                                  _          _ _              _                    _
    |___ \   _ __   ___  ___  ___ _ __ ___| |__   ___| | |   __ _  ___| |_   _ __ ___   ___| |_ __ _
      __) | | `_ \ / _ \/ __|/ _ \ `__/ __| `_ \ / _ \ | |  / _` |/ _ \ __| | `_ ` _ \ / _ \ __/ _` |
     / __/  | |_) | (_) \__ \  __/ |  \__ \ | | |  __/ | | | (_| |  __/ |_  | | | | | |  __/ || (_| |
    |_____| | .__/ \___/|___/\___|_|  |___/_| |_|\___|_|_|  \__, |\___|\__| |_| |_| |_|\___|\__\__,_|
            |_|                                             |___/
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utl_submit_ps64('
    Get-ItemProperty -Path d:\xls\have.xlsx | Format-List -Property * | clip;
    ');

    filename clp clipbrd;
    data want ;
       length var $20 val $60;
       infile clp;
       input;
       if _infile_=: " " then delete;;
       putlog _infile_;
       var=scan(_infile_,1,':');
       val=compbl(scan(_infile_,2,':'));
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*  WANT total obs=30                                                                                                     */
    /*                                                                                                                        */
    /*   VAR                  VAL                                                                                             */
    /*                                                                                                                        */
    /*   PSPath               Microsoft.PowerShell.Core\FileSystem                                                            */
    /*   PSParentPath         Microsoft.PowerShell.Core\FileSystem                                                            */
    /*   PSChildName          have.xlsx                                                                                       */
    /*   PSDrive              D                                                                                               */
    /*   PSProvider           Microsoft.PowerShell.Core\FileSystem                                                            */
    /*   Mode                 -a----                                                                                          */
    /*   VersionInfo          File                                                                                            */
    /*   BaseName             have                                                                                            */
    /*   Target               {}                                                                                              */
    /*   LinkType                                                                                                             */
    /*   Name                 have.xlsx                                                                                       */
    /*   Length               9572                                                                                            */
    /*   DirectoryName        D                                                                                               */
    /*   Directory            D                                                                                               */
    /*   IsReadOnly           False                                                                                           */
    /*   Exists               True                                                                                            */
    /*   FullName             D                                                                                               */
    /*   Extension            .xlsx                                                                                           */
    /*   CreationTime         3/29/2024 4                                                                                     */
    /*   CreationTimeUtc      3/29/2024 11                                                                                    */
    /*   LastAccessTime       4/2/2024 1                                                                                      */
    /*   LastAccessTimeUtc    4/2/2024 8                                                                                      */
    /*   LastWriteTime        3/31/2024 11                                                                                    */
    /*   LastWriteTimeUtc     3/31/2024 6                                                                                     */
    /*   Attributes           Archive                                                                                         */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */










































/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

/*----  output is in the widoes paste buffer -just paste                 ----*/


PSPath            : Microsoft.PowerShell.Core\FileSystem::D:\xls\have.xlsx
PSParentPath      : Microsoft.PowerShell.Core\FileSystem::D:\xls
PSChildName       : have.xlsx
PSDrive           : D
PSProvider        : Microsoft.PowerShell.Core\FileSystem
Mode              : -a----
VersionInfo       : File:             D:\xls\have.xlsx
                    InternalName:
                    OriginalFilename:
                    FileVersion:
                    FileDescription:
                    Product:
                    ProductVersion:
                    Debug:            False
                    Patched:          False
                    PreRelease:       False
                    PrivateBuild:     False
                    SpecialBuild:     False
                    Language:

BaseName          : have
Target            : {}
LinkType          :
Name              : have.xlsx
Length            : 9572
DirectoryName     : D:\xls
Directory         : D:\xls
IsReadOnly        : False
Exists            : True
FullName          : D:\xls\have.xlsx
Extension         : .xlsx
CreationTime      : 3/29/2024 4:24:04 PM
CreationTimeUtc   : 3/29/2024 11:24:04 PM
LastAccessTime    : 4/2/2024 12:27:16 PM
LastAccessTimeUtc : 4/2/2024 7:27:16 PM
LastWriteTime     : 3/31/2024 11:16:20 AM
LastWriteTimeUtc  : 3/31/2024 6:16:20 PM
Attributes        : Archive

















SAS Populates the details window.
The detail window may not show Auther but it is available


 Right click on file select properties > details

 _________________________________________________________
|                                                        |
|                                                        |
|  Right click on file select properties > details       |
|                                                        |
|  Roger.xlsx Properties                                 |
|                                                        |
|   +-------------------------------------------------+  |
|   | General| Security | Details | Previous Versions |  |
|   +-------------------------------------------------+  |
|                                                        |
|     Property    Value                                  |
|                                                        |
|     Discription ---------------------------------------|
|                                                        |
|     Title:      Demographics for Target Protocol       |
|     Subject:    Pubs                                   |
|     Tags:       SDTM clinical                          |
|     Categories  Demographics                           |
|     Comments:   Rogers Comments                        |
|     Author:     Roger Deangelis                        |
|                                                        |
|________________________________________________________|




%utl_submit_pl64('
use strict;`
use Win32::OLE;`
my $file = "d:\\xls\\roger.xlsx";
my $obj = Win32::OLE->GetObject($file);`
my $title = $obj->{"Title"};`
my $subject = $obj->{"Subject"};`
my $tags = $obj->{"Keywords"};`
my $category = $obj->{"Categories"};`
my $comments = $obj->{"Comments"};`
my $author = $obj->{"Author"};`
my $name = $obj->{"Name"};`

open (FH, ">", "d:/txt/roger.txt");`
print FH "File: $file\n";`
print FH "Author:$author\n";`
print FH "Title: $title\n";`
print FH "Subject: $subject\n";`
print FH "Tags:$tags\n";`
print FH "Category:$category\n";`
print FH "Comments: $comments\n";`
print FH "Name: $name\n";`
');

data want;
 infile "d:/txt/roger.txt";
 input;
 put _infile_;
 Property = scan(_infile_,1,':');
 Value    = scan(_infile_,2,':');
run;quit;



 Name
Date created             ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;

Author
Product name
Whether the file is a .NET file or not
Media tags for common media file formats
File metadata
Extended image information (e.g. ISO, brightness, aperture)
Title
Subject
Rating
Tags
Categories
Album
Genre
Item name display
Item type text
File version
Item folder path display
File size
Date created
Date modified
File attributes
File owner
Offline availability























/*----                                                                   ----*/


























data want;
 infile "d:/txt/roger.txt";
 input;
 put _infile_;
 Property = scan(_infile_,1,':');
 Value    = scan(_infile_,2,':');
run;quit;


run;quit;run;quit;











































%utl_submit_pl64('
use strict;`
use Win32::OLE;`
open (FH, ">", "d:/txt/roger.txt");`
my $file = "d:\\xls\\roger.xlsx";`
print FH "File: $file\n";`
my $obj = Win32::OLE->GetObject($file);`
my $subject = $obj->{"Subject"};`
print FH "Subject: Subject:subject\n";`
my $title = $obj->{"Title"};`
print FH "Title:$title\n";`
my $comments = $obj->{"Comments"};`
print FH "Comments $comments\n";`
');


my $concat =  $subject . "@" . $title . "@" . $comments;`
print FH "$concat\n";`
data _null_;
  rc = filename('tmp', 'd:/xls/roger.xlsx');
  fid = fopen('tmp');
  if fid then do;
    infonum = foptnum(fid);
    do i=1 to infonum;
      infoname = foptname(fid,i);
      res=finfo(fid, infoname);
      put infoname= res=;
    end;
  end;
  rc = fclose(fid);
run;

%utl_submit_ps64('
Get-ItemProperty -Path d:\xls\have.xlsx | Format-List -Property * | clip;
');

filename clp clipbrd;

data lower_panel;
   length var $20 val $60;
   infile clp;
   input;
   if _infile_=: " " then delete;;
   putlog _infile_;
   var=scan(_infile_,1,':');
   val=compbl(scan(_infile_,2,':'));
run;quit;

right click on dd:/xls/

roger.x



what is the python syntax to read subject in the windows file properties details window

/*----                                                                   ----*/
/*---- note the backtic is used to separate lines                        ----*/
/*----                                                                   ----*/

%utl_submit_pl64('
use strict;`
use Win32::OLE;`
open (FH, ">", "d:/txt/roger.txt");`
print FH "Hello, World!\n";`
my $txt = "roger\n";`
print fh $txt;`
my $file = "d:\\xls\\roger.xlsx";`
print FH $file;`
my $obj = Win32::OLE->GetObject($file);`
print FH "subject\n";`
my $subject = $obj->{"Subject"};`
print FH "$subject\n";`
my $title = $obj->{"Title"};`
print FH "$title\n";`
my $comments = $obj->{"Comments"};`
print FH "$comments\n";`
my $concat =  $subject . "@" . $title . "@" . $comments;`
print FH "$concat\n";`
');


%utl_submit_pl64('
use strict;`
use Win32::OLE;`
open (FH, ">", "d:/txt/roger.txt");`
my $file = "d:\\xls\\roger.xlsx";`
my $obj = Win32::OLE->GetObject($file);`
print FH "subject\n";`
my $subject = $obj->{"Subject"};`
print FH "$subject\n";`
my $title = $obj->{"Title"};`
print FH "$title\n";`
my $comments = $obj->{"Comments"};`
print FH "$comments\n";`
my $concat =  $subject . "@" . $title . "@" . $comments;`
print FH "$concat\n";`
');


%utl_submit_pl64('
use strict;
open (FH, ‘>’, 'd:/txt/roger.txt”);
my $file = "d:\\xls\\roger.xlsx";`
my $metadata = File::Metadata->new($file);
my $origin = $metadata->get("origin");`
print "$origin";
');

# Get the metadata
my $origin = $metadata->get('origin');



    ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;








my $all_metadata = $metadata->get_all();`
print $all_metadata;`
');















my $width = $all_metadata->{width};
my $height = $all_metadata->{height};



use File::Metadata;





















Clipboard->copy("TEXT");`
filename clp clipbrd;
data want;
 infile;
 subject = scan(_infile_,1,'@');
 title   = scan(_infile_,2,'@');
 comments =

filename clp clipbrd;
use Clipboard;`
use Xclip;`
Clipboard->copy("Test");`




















                            ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;

%utl_submit_pl64('
use Win32::OLE;`
use File::Basename;`
my $file = "d:\xls\roger.xlsx";`
my $shell = Win32::OLE->new("Shell.Application");`
my $folder = $shell->NameSpace(File::Basename::dirname($file));`
my $fileObj = $folder->ParseName(File::Basename::basename($file));`
my $subject = $fileObj->GetDetailsOf(0, 21);`
print "Subject: $subject\n";`
');


%utl_submit_pl64('
use Win32::OLE;`
$file = "d:\\xls\\roger.xlsx";`
print "$file\n";`
$shell = Win32::OLE->new("Shell.Application");`
$folder = $shell->NameSpace(File::Basename::dirname("d:\\xls\\roger.xlsx"));`
print "$folder\n";`
');
$fileObj = $folder->ParseName(File::Basename::basename($file));`
print "$fileObj\n";`
$subject = $fileObj->GetDetailsOf(0, 21);`
print "Subject: $subject\n";`
');


use Win32::OLE;

my $file = "C:\path\to\your\file.txt";

my $shell = Win32::OLE->new('Shell.Application');
my $folder = $shell->NameSpace(File::Basename::dirname($file));
my $fileObj = $folder->ParseName(File::Basename::basename($file));

# Retrieve various file properties
my $title = $fileObj->GetDetailsOf(0, 2);
my $subject = $fileObj->GetDetailsOf(0, 21);
my $tags = $fileObj->GetDetailsOf(0, 26);
my $comments = $fileObj->GetDetailsOf(0, 24);

print "Title: $title\n";
print "Subject: $subject\n";
print "Tags: $tags\n";
print "Comments: $comments\n";

The key steps are:






























#!/usr/bin/perl`
$txt = "roger"`
print $txt`
');

use Win32::OLE`
my $file = 'd:\\xls\\roger.xlsx'`
my $obj = Win32::OLE->GetObject(my $file)`
my $subject = $obj->{'Subject'}`
print my $subject`
");

                ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;
print "File subject: $subject\n";

%utl_pybegin;
parmcards4;
import pywin32
ver_parser = Dispatch("Shell.Application");
file_obj = ver_parser.NameSpace("d:\\xls\\roger.xlsx");
file_properties = file_obj.GetDetailsOf(file_obj.ParseName("roger.xlsx"), 0)`
print("File properties:", file_properties);
;;;;
%utl_pyend;





          ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;



perl -MCPAN -e Win32::OLE













PSPath            : Microsoft.PowerShell.Core\FileSystem::D:\xls\have.xlsx
PSParentPath      : Microsoft.PowerShell.Core\FileSystem::D:\xls
PSChildName       : have.xlsx
PSDrive           : D
PSProvider        : Microsoft.PowerShell.Core\FileSystem
Mode              : -a----
VersionInfo       : File:             D:\xls\have.xlsx
                    InternalName:
                    OriginalFilename:
                    FileVersion:
                    FileDescription:
                    Product:
                    ProductVersion:
                    Debug:            False
                    Patched:          False
                    PreRelease:       False
                    PrivateBuild:     False
                    SpecialBuild:     False
                    Language:

BaseName          : have
Target            : {}
LinkType          :
Name              : have.xlsx
Length            : 9572
DirectoryName     : D:\xls
Directory         : D:\xls
IsReadOnly        : False
Exists            : True
FullName          : D:\xls\have.xlsx
Extension         : .xlsx
CreationTime      : 3/29/2024 4:24:04 PM
CreationTimeUtc   : 3/29/2024 11:24:04 PM
LastAccessTime    : 4/1/2024 2:17:41 PM
LastAccessTimeUtc : 4/1/2024 9:17:41 PM
LastWriteTime     : 3/31/2024 11:16:20 AM
LastWriteTimeUtc  : 3/31/2024 6:16:20 PM
Attributes        : Archive


what is the perl syntax to extract origin data fron file properties detail















Table 1.1 Demographics for Target Protocol


















/* T1002960

Author of excel file importing to SAS

inspired by
https://goo.gl/UbGFGd
https://communities.sas.com/t5/Base-SAS-Programming/author-of-excel-file-importing-to-SAS/m-p/338811

Extracting file meta data Owner, dates, author, title

* you can set properties but you need microsoft only tools to extract the meta data?
see microsoft only script below




HAVE

  c:\utl\utl_toc_xls.sas

WANT
====

SOLUTION 1

NFOCNT=6
information for a      Sequential File:
filename           c:\utl\\utl_toc_xls.sas
RECFM              V
LRECL              384
File Size (bytes)  20755
Last Modified      07Mar2017:05:35:28
Create Time        06Mar2017:08:23:43

SOLUTION 2

FILE_OWNER=workstation
FILE_NAME=utl_toc_xls.sas
FILE_DATE=2017-03-07
FILE_TIME=5:35:00
FILE_SIZE=20755

SOLUTION 3

File          c:\utl\utl_toc_xls.sas
File size     20755 bytes
Created       06Mar2017:08:23:43
Last modified 07Mar2017:05:35:28

MS only solution
================

All the MS office meta data available,
not all meta data available for all MS applications.

Title
Subject
Author
Keywords
Comments
Template
Last author
Revision number
Application name
Last print date
Creation date
Last save time
Total editing time
Number of pages
Number of words
Number of characters
Security
Category
Format
Manager
Company
Number of bytes
Number of lines
Number of paragraphs
Number of slides
Number of notes
Number of hidden Slides
Number of multimedia clips
Hyperlink base
Number of characters (with spaces)

EXCEL

Title - Metadata Test
Subject - Excel Test Scripts
Author - Ken Myer
Keywords - testing, scripts
Comments - This is a sample spreadsheet used for testing purposes.
Template -
Last author - Ken Myer
Revision number -
Application name - Microsoft Excel
Creation date - 6/13/2006 8:40:17 PM
Last save time - 6/13/2006 9:07:15 AM
Security - 0
Category -
Format -
Manager -
Company - Microsoft Corporation
Hyperlink base -


On Error Resume Next

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")

For Each strProperty in objWorkbook.BuiltInDocumentProperties
    Wscript.Echo strProperty.Name & " - " & strProperty.Value
Next


SOLUTION 1


filename x pipe 'dir /q c:\utl\utl_toc_xls.sas';
data want;
   retain file_owner file_name;
   infile x firstobs=6 truncover;
   input @1 file_date mmddyy10.
        @13 file_time time8.
            file_size comma19.
            file_owner $22.
            file_name $32.;
   list;
   format file_date yymmdd10. file_time time8.;
   put
     file_owner=  /
     file_name =  /
     file_date =  /
     file_time =  /
     file_size =  /
   ;
   stop;
run;
filename x clear;

RULE:     ----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----0
6         03/07/2017  05:35 AM            20,755 backup-PC\backup       utl_toc_xls.sas 77

SOLUTION 2

%utlopts;
%macro FileAttribs(filename);
   %global rc fid fidc;
   %global Bytes CreateDT ModifyDT;
   %let rc=%sysfunc(filename(onefile,&filename));
   %let fid=%sysfunc(fopen(&onefile));
   %let Bytes=%sysfunc(finfo(&fid,File Size (bytes)));
   %let CreateDT=%qsysfunc(finfo(&fid,Create Time));
   %let ModifyDT=%qsysfunc(finfo(&fid,Last Modified));
   %let fidc=%sysfunc(fclose(&fid));
   %let rc=%sysfunc(filename(onefile));
   %put  File &filename ;
   %put  File size is &Bytes bytes;
   %put  Created &CreateDT;
   %put  Last modified &ModifyDT;
%mend FileAttribs;

%FileAttribs (c:\utl\utl_toc_xls.sas) ;

SOLUTION 3

File          c:\utl\utl_toc_xls.sas
File size     20755 bytes
Created       06Mar2017:08:23:43
Last modified 07Mar2017:05:35:28
%utlnopts;
data _null_;
   length opt $100 optval $100;

   /* Allocate file  */
   rc=FILENAME('myfile',
      'c:\utl\\utl_toc_xls.sas');

   /* Open file */
   fid=FOPEN('myfile');

   /* Get number of information
      items */
   infocnt=FOPTNUM(fid);
   put infocnt=;

   /* Retrieve information items
      and print to log  */
   put @1 'Information for a
      Sequential File:';
   do j=1 to infocnt;
      opt=FOPTNAME(fid,j);
      optval=FINFO(fid,upcase(opt));
      put @1 opt @20 optval;
   end;

   /* Close the file */
   rc=FCLOSE(fid);

   /* Deallocate the file */
   rc=FILENAME('myfile');
run;

               ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;
data info;
   length infoname infoval $60;
   drop rc fid infonum i close;
   rc=filename('abc','d:\xls\roger.xlsx');
   fid=fopen('abc');
   infonum=foptnum(fid);
   do i=1 to infonum;
      infoname=foptname(fid,i);
      infoval=finfo(fid,infoname);
      output;
   end;
   close=fclose(fid);
run;










Get-FileMetaData -File "C:\path\to\file.txt"

# Get metadata for multiple files
$files = "C:\path\to\file1.txt", "C:\path\to\file2.exe"
Get-FileMetaData -File $files

# Get metadata for all files in a directory
Get-ChildItem -Path "C:\path\to\directory" -Force | Get-FileMetaData

# Get metadata and include the file's digital signature
Get-FileMetaData -File "C:\path\to\file.exe" -Signature



Get-ItemProperty -Path 'd:\xls\SDTMTerminology.xlsx'

Get-FileMetaData -File "d:\xls\SDTMTerminology.xlsx" | clip


Set-ItemProperty -Path <file_path> -Name <property_name> -Value <new_value>




 d:\xls\SDTMTerminology.xlsx

Get-ChildItem 'd:\xls\SDTMTerminology.xlsxt' | Get-Member -MemberType Property


 Get-FileMetaData -File "d:\xls\SDTMTerminology.xlsx" | clip

Get-ItemProperty -Path 'd:\xls\SDTMTerminology.xlsx'
foreach-Object{
    $_.SubItems[0].Text + " , " + $_.SubItems[1].Text + " , " + $_.SubItems[2].Text
} | clip



     Get-ItemProperty -Path 'd:\xls\SDTMTerminology.xlsx'
     foreach-Object{
         $_.SubItems[0].Text + " , " + $_.SubItems[1].Text + " , " + $_.SubItems[2].Text
     } | clip




$file = Get-ChildItem 'd:\xls\SDTMTerminology.xlsx'
$file.Name
$file.Length
$file.CreationTime
$file.LastWriteTime



Get-FileMetaData -File "d:\xls\have.xlsx" | clip


$File = Get-Item -Path "d:\xls\have.xlsx"
$File.ListItemAllFields


Get-ItemProperty -Path d:\xls\have.xlsx | Format-List -Property * | clip
(Get-ItemProperty -Path d:\xls\roger.xlsx).Authors | Format-List -Property * | clip
OWNER='johndoe'
OWNER='johndoe'





PSPath            : Microsoft.PowerShell.Core\FileSystem::D:\xls\have.xlsx
PSParentPath      : Microsoft.PowerShell.Core\FileSystem::D:\xls
PSChildName       : have.xlsx
PSDrive           : D
PSProvider        : Microsoft.PowerShell.Core\FileSystem
Mode              : -a----
VersionInfo       : File:             D:\xls\have.xlsx
                    InternalName:
                    OriginalFilename:
                    FileVersion:
                    FileDescription:
                    Product:
                    ProductVersion:
                    Debug:            False
                    Patched:          False
                    PreRelease:       False
                    PrivateBuild:     False
                    SpecialBuild:     False
                    Language:

BaseName          : have
Target            : {}
LinkType          :
Name              : have.xlsx
Length            : 9572
DirectoryName     : D:\xls
Directory         : D:\xls
IsReadOnly        : False
Exists            : True
FullName          : D:\xls\have.xlsx
Extension         : .xlsx
CreationTime      : 3/29/2024 4:24:04 PM
CreationTimeUtc   : 3/29/2024 11:24:04 PM
LastAccessTime    : 4/1/2024 12:01:29 PM
LastAccessTimeUtc : 4/1/2024 7:01:29 PM
LastWriteTime     : 3/31/2024 11:16:20 AM
LastWriteTimeUtc  : 3/31/2024 6:16:20 PM
Attributes        : Archive


FILENAME fileref 'external-file'
   <ENCODING='encoding-value'>
   <TERMSTR='CRLF'|'LF'|'LFCR'>
   <LRECL=record-length>
   <RECFM=record-format>
   <BLKSIZE=block-size>
   <DISP=('STATUS','DISPOSITION')>
   <OWNER=owner-id>
   <PERMS=access-permissions>
   <OPMODES=open-mode>
   <FILEEXT=file-extension>
   <FILETYPE=file-type>
   <XMLTYPE=xml-type>
   <XMLMAP=xml-map-file>
   <XMLSCHEMA=xml-schema-file>
   <XMLCLIENT=xml-client-file>
   <XMLDBMS=xml-dbms-file>
   <XMLDBTYPE=xml-dbms-type>
   <XMLDOCTYPE=xml-doctype>
   <XMLENCODING=xml-encoding>
   <XMLMAP=xml-map-file>
   <XMLNAMESPACE=xml-namespace>
   <XMLROOT=xml-root>
   <XMLSCHEMA=xml-schema-file>
   <XMLVERSION=xml-version>
   <XMLXSL=xml-xsl-file>;




Set-ItemProperty -Path "D:\xls\have.xlsx" -Name Author -Value "Roger DeAngelis" | Format-List -Property * | clip   ;






$excelFile = "D:\xls\have.xlsx"
$author = "John Doe"
Set-ItemProperty -Path $excelFile -Name Author -Value $author








$FilePath = "D:\xls\have.xlsx"
 echo ("The file "+ $FilePath +" has been processed.")
$ShellApplication = New-Object -ComObject Shell.Application
$ShellFolder = $ShellApplication.NameSpace($FilePath)
$MetadataProperties = [ordered]@{}
$ShellFolder.Items() | ForEach-Object {
    $DataValue = $ShellFolder.GetDetailsOf($_, 0)
    $PropertyValue = (Get-Culture).TextInfo.ToTitleCase($DataValue.Trim()).Replace(' ', '')
    echo ("The file "+ $PropertyValue +" has been processed.")
    if ($PropertyValue -ne '') {
        $MetadataProperties["Description"] = $PropertyValue
    }
}






echo ("The file "+ $PropertyValue +" has been processed.")





PSPath            : Microsoft.PowerShell.Core\FileSystem::D:\xls\have.xlsx
PSParentPath      : Microsoft.PowerShell.Core\FileSystem::D:\xls
PSChildName       : have.xlsx
PSDrive           : D
PSProvider        : Microsoft.PowerShell.Core\FileSystem
Mode              : -a----
VersionInfo       : File:             D:\xls\have.xlsx
                    InternalName:
                    OriginalFilename:
                    FileVersion:
                    FileDescription:
                    Product:
                    ProductVersion:
                    Debug:            False
                    Patched:          False
                    PreRelease:       False
                    PrivateBuild:     False
                    SpecialBuild:     False
                    Language:

BaseName          : have
Target            : {}
LinkType          :
Name              : have.xlsx
Length            : 9572
DirectoryName     : D:\xls
Directory         : D:\xls
IsReadOnly        : False
Exists            : True
FullName          : D:\xls\have.xlsx
Extension         : .xlsx
CreationTime      : 3/29/2024 4:24:04 PM
CreationTimeUtc   : 3/29/2024 11:24:04 PM
LastAccessTime    : 4/1/2024 2:12:12 PM
LastAccessTimeUtc : 4/1/2024 9:12:12 PM
LastWriteTime     : 3/31/2024 11:16:20 AM
LastWriteTimeUtc  : 3/31/2024 6:16:20 PM
Attributes        : Archive
Write and read meta data saved in the windows file properties details panel  
