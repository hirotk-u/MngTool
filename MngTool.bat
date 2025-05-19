@echo off
rem ###########################################################################
rem # Management tool
rem ###########################################################################
set excel_file_path1=D:\workspace_bat\MngTool\TestFolder\TestFile1.xlsx
set excel_file_path2=D:\workspace_bat\MngTool\TestFolder\TestFile2.xlsx


rem ==================================
rem åªç›ì˙éûäiî[óp(ïœçXÇµÇ»Ç¢Ç±Ç∆)
rem ==================================
set year=%date:~0,4%
set month=%date:~5,2%
set day=%date:~8,2%
set hh=%time:~0,2%
set mm=%time:~3,2%
set ss=%time:~6,2%

rem åªç›ì˙ït
set now_datetime=%year%%month%%day

rem ==================================
rem Main
rem ==================================
echo *****************************
echo Menu
echo 1.Open excel file1
echo 2.Open excel file2
echo 9:End
echo *****************************

set /p NUM1="Input number:"
if %NUM1%==1 goto 1
if %NUM1%==2 goto 2
if %NUM1%==9 goto exit_label

set select_num=0
rem 1 --------------------
:1
set select_num=1
start /wait excel %excel_file_path1%
goto result_label

rem 2 --------------------
:2
set select_num=2
start /wait excel %excel_file_path2%
goto result_label

rem --- result ---------------------------
:result_label
echo.
set /p selected="Completed input? (y or n):"
if %selected%==y goto :yes
if %selected%==y goto :no


rem --- no ----------------------------------------
:no
echo "No result writing"
goto end_label

rem --- yes ----------------------------------------
:yes
echo "Write result"

rem TODO write proc

goto end_label

rem --- End ------------------------------
:end_label
pause

