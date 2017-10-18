Import xlsm to sas dataset

  WORKING CODE (Pretty much all of it)

     class <- read.xlsx("d:/xls/class_final.xlsm",sheetName="class",as.data.frame=TRUE);
     import r=class  data=wrk.class;   * save as class.sas7bdat in parent work directory;


For input xlsm workbook see
https://www.dropbox.com/s/tokxyqz4p292z4h/class_final.xlsm?dl=0

see
https://goo.gl/dLnTpP
https://communities.sas.com/t5/SAS-Procedures/Problem-with-libname-excel-to-import-xlsm-file/m-p/405226


HAVE (simple python code exists to open workbook and execute macro(if it was not executed on open) - not provided)
===================================================================================================================

   https://www.dropbox.com/s/tokxyqz4p292z4h/class_final.xlsm?dl=0
   which I saved locally as d:/xls/class_final.xlsm

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+
   21 |            |            |            |            |  1900.9    | Macro supplied total
      +------------+------------+------------+------------+------------+ on opening workbook

      [CLASS]


WANT
====

WORK.CLAA total obs=20

Obs    OBS    NAME       SEX    AGE    HEIGHT    WEIGHT

  1      0    Alfred      M      14     69.0      112.5
  2      1    Alice       F      13     56.5       84.0
  3      2    Barbara     F      13     65.3       98.0
  4      3    Carol       F      14     62.8      102.5
  5      4    Henry       M      14     63.5      102.5
  6      5    James       M      12     57.3       83.0
  7      6    Jane        F      12     59.8       84.5
  8      7    Janet       F      15     62.5      112.5
  9      8    Jeffrey     M      13     62.5       84.0
 10      9    John        M      12     59.0       99.5
 11     10    Joyce       F      11     51.3       50.5
 12     11    Judy        F      14     64.3       90.0
 13     12    Louise      F      12     56.3       77.0
 14     13    Mary        F      15     66.5      112.0
 15     14    Philip      M      16     72.0      150.0
 16     15    Robert      M      12     64.8      128.0
 17     16    Ronald      M      15     67.0      133.0
 18     17    Thomas      M      11     57.5       85.0
 19     18    William     M      15     66.5      112.0
 20      .                        .       .      1900.5   * supplied by macro


*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;

   https://www.dropbox.com/s/tokxyqz4p292z4h/class_final.xlsm?dl=0
   which I saved locally as d:/xls/class_final.xlsm

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;


%utl_submit_wps64('
options set=R_HOME "C:/Program Files/R/R-3.3.2";
libname wrk sas7bdat "%sysfunc(pathname(work))";
proc r;
submit;
source("C:/Program Files/R/R-3.3.2/etc/Rprofile.site", echo=T);
library(xlsx);
class <- read.xlsx("d:/xls/class_final.xlsm",sheetName="class",as.data.frame=TRUE);
endsubmit;
import r=class  data=wrk.class;
run;quit;
');

*                       _                     __       _ _
__      ___ __  ___    (_) __ ___   ____ _   / _| __ _(_) |
\ \ /\ / / '_ \/ __|   | |/ _` \ \ / / _` | | |_ / _` | | |
 \ V  V /| |_) \__ \   | | (_| |\ V / (_| | |  _| (_| | | |
  \_/\_/ | .__/|___/  _/ |\__,_| \_/ \__,_| |_|  \__,_|_|_|
         |_|         |__/
;

%utl_submit_r64('
source("C:/Program Files/R/R-3.3.2/etc/Rprofile.site", echo=T);
library(xlsx);
class <- read.xlsx("d:/xls/class_final.xlsm",sheetName="class",as.data.frame=TRUE);
class;
save(class,file="d:/rda/class.rda", compress = FALSE);
load(file = "d:/rda/class.rda");
class;
');


%utl_submit_wps64('
options set=R_HOME "C:/Program Files/R/R-3.3.2";
libname wrk sas7bdat "%sysfunc(pathname(work))";
proc r;
submit;
source("C:/Program Files/R/R-3.3.2/etc/Rprofile.site", echo=T);
load(file = "d:/rda/class.rda");
class;
endsubmit;
import r=class  data=wrk.class;
run;quit;
');
