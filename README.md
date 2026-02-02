# frontal_impact_triaxial_acceleration

The Analysis folder has three files: 'Analysis by vehicle type & segment' and 'Stats'. 
'Analysis by vehicle type & segment' is a summary of all IIHS frontal crash test data from vehicles with model years 1997-2025, including
- vehicle info (year, make, model)
- vehicle type (car, SUV, minivan, pickup)
- vehicle delta-v (kph)
- vehicle peak acceleration [x, y, z, resultant] (g)
- driver ATD head and chest peak acceleration [x, y, z, resultant] (g)
- passenger ATD peak head, chest, and pelvis acceleration [x, y, z, resultant] (g)
- resultant ATD-vehicle acceleration ratios
- directional-resultant ATD acceleration ratios
The data is separated by vehicle type in different sheets ('Car', 'SUV', 'Minivan', 'Pickup'), as well as compiled together ('All Data')
The 'Plots' sheets contains graphs comparing the ATD-vehicle acceleration ratios and directional-resultant acceleration ratios
The 'Year range' sheets lists all vehicle info from all frontal tests

'Stats' contains one-way ANOVA testing and Bonferroni-corrected post-hoc tests when significant differences were observed for ATD-vehicle and directional-resultant acceleration ratios

'peak acceleration timing stats' contains peak accelerations and times to peak acceleration for vehicle, driver head and chest, and passenger head, chest, and pelvis in the sheet 'All Timing & Accel Data' and one-way ANOVA testing and Bonferroni-corrected post-hoc tests when significant differences were observed for the peak acceleration timings in the sheet 'a Timing Stats - by location'

The 'Data Download & Analysis Code' has four folders: 'extract data', 'file download', 'stats', and 'write to excel'.
'extract data' contains MATLAB files to get data from IIHS frontal crash test excels
'file download' contains python files to download IIHS frontal crash test excels from the IIHS TestData database
'stats' contains a MATLAB file confirming the ANOVA testing in 'Stats'
'write data to excel' takes the extracted data returned from the MATLAB files in 'extract data' and converts it to excel files
