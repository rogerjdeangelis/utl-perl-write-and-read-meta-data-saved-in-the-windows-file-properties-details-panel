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
