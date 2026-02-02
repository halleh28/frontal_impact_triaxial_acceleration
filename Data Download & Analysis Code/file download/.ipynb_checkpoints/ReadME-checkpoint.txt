iihs_data_download_frontal_original.py downloads original frontal impact test excels from IIHS database
*starts at pg 3 - remove CF97018, CF98004, CF98006, CF98001, CF97024, CF98005, CF97021, CF97026, CF98003, CF97023, CF97022, CF98014, CF98009, CF98010, CF98002, CF99017 (intrusion data, not accel data)
*ends at pg 16 - remove CEF1608, CEF1609, CF18004, CF19001, CF19002 (no accel data in excels)
iihs_data_download_frontal_original_diadem.py downloads .tdms (diadem) files for original frontal impact tests with no accel data in summary excels from IIHS database
*starts at pg 17 - remove CEF1602 and CEF1603 (accel data in excel) and CEF2212 (no vehicle y accel data, resulting in actual vehicle resultant accel lower than reported)
*download CEF1608, CEF1609, CF18004, CF19001, CF19002 from pg 16
iihs_data_download_frontal_updated.py downloads all updated frontal impact tests .tdms (diadem) files from IIHS database
*remove CEF2348 (no vehicle accel data)
*must input data page tabs to run between

User must download python packages: selenium, pyautogui (using Anaconda instructions, run Anaconda prompt in administrator mode)

