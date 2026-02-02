load('IIHS_diadem_d','frontal_info')

header = {'Test ID' 'Vehicle Info' 'Vehicle dV (kph)' 'Vehicle Peak a (g)' 'Driver Head Peak a - x (g)' 'Driver Head Peak a - y (g)' 'Driver Head Peak a - z (g)' 'Driver Head Peak a - R (g)', 'Driver Chest Peak a - x (g)' 'Driver Chest Peak a - y (g)' 'Driver Chest Peak a - z (g)' 'Driver Chest Peak a - R (g)'};
xlswrite('Results_diadem_d.xls', header, 'Sheet1', 'A1')
xlswrite('Results_diadem_d.xls', frontal_info, 'Sheet1', 'A2')
close all