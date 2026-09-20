REM @ECHO OFF
ECHO XLS Merge %1 %2
REM Sintassi file_size_analyzer path file_csv_folders
REM Salva nel file file_size.csv il contenuto delle cartelle e path preciso.
REM Se file_csv_folders è uguale a #, estrae tutto il contenuto della cartella, non solo la lista passata
for %%F in (%0) do set PYSCRIPT=ntj_fact.py
SET PYN=k:\Tools\pyn.cmd
"%PYN%" "%PYSCRIPT%" %1 %2
