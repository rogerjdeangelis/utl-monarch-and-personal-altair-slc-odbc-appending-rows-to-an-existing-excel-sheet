# utl-monarch-and-personal-altair-slc-odbc-appending-rows-to-an-existing-excel-sheet
Monarch and personal altair slc ODBC appending rows to an existing excel sheet
    %let pgm=utl-monarch-and-personal-altair-slc-odbc-appending-rows-to-an-existing-excel-sheet;

    %stop_submission;

    Monarch and personal altair slc ODBC appending rows to an existing excel sheet

    github
    https://github.com/rogerjdeangelis/utl-monarch-and-personal-altair-slc-odbc-appending-rows-to-an-existing-excel-sheet

    OPS PROBLEM
    I guess you're appending data to a worksheet in Excel? Maybe someone changed
    the name of the worksheet? Or if you're using named ranges, perhaps a user
    has inadvertently changed the name or shape of these?

    SOLUTION

      Personal altair slc ODBC works (should be able to use passthru to sql or slc proc sql)

      Simple solution using, scan_text=no, does not work.
      4416      libname xel "d:/xls/class.xlsx" scan_text=no;
      ERROR: Option "scan_text" is not known for the LIBNAME statement

    Monarch post
    https://community.altair.com/discussion/comment/187055?tab=all#Comment_187055?utm_source=community-search&utm_medium=organic-search&utm_term=monarch+excel

    see for scan_text=no solution
    https://github.com/rogerjdeangelis/utl-appending-records-to-an-existing-excel-sheet

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    d:/xls/class.xlsx

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
     1| NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
     2| BASE       |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
     3| BASE       |    F       |    15      |   56.5     |   84       |
      +------------+------------+------------+------------+------------+

     [BASE]

    TABLE MONDAY (data to append)

     NAME     SEX    AGE    HEIGHT    WEIGHT

    MONDAY     F      13     65.3       98.0
    MONDAY     F      14     62.8      102.5

    /*--- create input ----*/
    %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    &_init_;

    %utlfkil(d:/xls/class.xlsx);  /*--- delete if exists ---*/

    libname xel excel "d:/xls/class.xlsx";

    data xel.base monday;
    informat
       NAME $8.
       SEX $1.
       AGE 8.
       HEIGHT 8.
       WEIGHT 8.
    ;
    input
       NAME SEX AGE HEIGHT WEIGHT;
     select;
       when (_n_<3) do;name="BASEDAY"; output xel.base;end;
       when (_n_<5) do;name="MONDAY "; output monday  ;end;
     end; /*-- leave off otherwise to force non inclusive error --*/
    cards4;
    Alfred M 14 69 112.5
    Alice F 13 56.5 84
    Barbara F 13 65.3 98
    Carol F 14 62.8 102.5
    ;;;;
    run;quit;
    libname xel clear;

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    &_init_;
    libname myexcel odbc
        noprompt="Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};
                   DBQ=d:\xls\class.xlsx;
                   ReadOnly=0;"
         ;

    /* Import CSV data */

    /* Append data to Excel sheet */
    proc append base=myexcel.'BASE$'n
                data=work.monday;
    run;

    /* Clear libname connection */
    libname myexcel clear;

    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */

    d:/xls/class.xlsx

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
     1| NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
     2| BASE       |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
     3| BASE       |    F       |    15      |   56.5     |   84       |
      +------------+------------+------------+------------+------------+ ---------
     4| MONDAY     |    F       |    13      |   65.3     |   98       |
      +------------+------------+------------+------------+------------+ added rows
     5|MONDAY      |    F       |    15      |   62.8     | 102,5      |
      +------------+------------+------------+------------+------------+ ---------

     [BASE]

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */
    5103      ODS _ALL_ CLOSE;
    5104      FILENAME WPSWBHTM TEMP;
    NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD9416\#LN00185
    5105      ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD9416";
    5106      %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    5107      &_init_;
    5108
    5109      %utlfkil(d:/xls/class.xlsx);  /*--- delete if exists ---*/
    5110
    5111      libname xel excel "d:/xls/class.xlsx";
    NOTE: Library xel assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/class.xlsx

    5112
    5113      data xel.base monday tuesday;
    5114      informat
    5115         NAME $8.
    5116         SEX $1.
    5117         AGE 8.
    5118         HEIGHT 8.
    5119         WEIGHT 8.
    5120      ;
    5121      input
    5122         NAME SEX AGE HEIGHT WEIGHT;
    5123       select;
    5124         when (_n_<3) do;name="BASEDAY"; output xel.base;end;
    5125         when (_n_<5) do;name="MONDAY "; output monday  ;end;
    5126         when (_n_<8) do;name="TUESDAY"; output tuesday ;end;
    5127       end; /*-- leave off otherwise to force non inclusive error --*/
    5128      cards4;

    NOTE: Data set "XEL.base" has an unknown number of observation(s) and 5 variable(s)
    NOTE: Data set "WORK.monday" has 2 observation(s) and 5 variable(s)
    NOTE: Data set "WORK.tuesday" has 2 observation(s) and 5 variable(s)
    NOTE: The data step took :
          real time : 0.257
          cpu time  : 0.062


    5129      Alfred M 14 69 112.5
    5130      Alice F 13 56.5 84
    5131      Barbara F 13 65.3 98
    5132      Carol F 14 62.8 102.5
    5133      Henry M 14 63.5 102.5
    5134      James M 12 57.3 83
    5135      ;;;;
    5136      run;quit;
    NOTE: Libref XEL has been deassigned.
    5137      libname xel clear;
    5138
    5139      libname myexcel odbc
    NOTE: The number of active statements per connection could not be determined. As a result, the connection type will be set to UNIQUE.
    NOTE: Library myexcel assigned as follows:
          Engine:        ODBC
          Physical Name:  (EXCEL version 12.00.0000)

    5140          noprompt=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    5141      XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    5142      XXXXXXXXXXXXXXXXXXXXXXXXXXX
    5143           ;
    5144
    5145      /* Import CSV data */
    5146
    5147      /* Append data to Excel sheet */
    5148      proc append base=myexcel.'BASE$'n
    5149                  data=work.monday;
    5150      run;
    NOTE: 2 observations were appended to data set "MYEXCEL.BASE$"
    NOTE: Procedure append step took :
          real time : 0.687
          cpu time  : 0.265


    NOTE: Libref MYEXCEL has been deassigned.
    5151
    5152      /* Clear libname connection */
    5153      libname myexcel clear;
    5154
    5155
    5156      quit; run;
    5157      ODS _ALL_ CLOSE;
    5158      FILENAME WPSWBHTM CLEAR;

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
