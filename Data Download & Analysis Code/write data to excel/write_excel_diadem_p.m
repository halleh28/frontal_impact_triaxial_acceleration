load('IIHS_diadem_p','frontal_info')

header = {'Test ID' 'Vehicle Info' 'Passenger Head Peak a - x (g)' 'Passenger Head Peak a - y (g)' 'Passenger Head Peak a - z (g)' 'Passenger Head Peak a - R (g)', 'Passenger Chest Peak a - x (g)' 'Passenger Chest Peak a - y (g)' 'Passenger Chest Peak a - z (g)' 'Passenger Chest Peak a - R (g)'};
xlswrite('Results_diadem_p.xls', header, 'Sheet1', 'A1')
xlswrite('Results_diadem_p.xls', frontal_info, 'Sheet1', 'A2')
close all
