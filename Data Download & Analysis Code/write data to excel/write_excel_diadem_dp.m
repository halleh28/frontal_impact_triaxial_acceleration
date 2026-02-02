load('IIHS_diadem_dp','frontal_info')

header = {'Test ID' 'Vehicle Info' 'Vehicle dV (kph)' 'Vehicle Peak a (g)' 'Driver Head Peak a - x (g)' 'Driver Head Peak a - y (g)' 'Driver Head Peak a - z (g)' 'Driver Head Peak a - R (g)', 'Driver Chest Peak a - x (g)' 'Driver Chest Peak a - y (g)' 'Driver Chest Peak a - z (g)' 'Driver Chest Peak a - R (g)', 'Passenger Head Peak a - x (g)' 'Passenger Head Peak a - y (g)' 'Passenger Head Peak a - z (g)' 'Passenger Head Peak a - R (g)', 'Passenger Chest Peak a - x (g)' 'Passenger Chest Peak a - y (g)' 'Passenger Chest Peak a - z (g)' 'Passenger Chest Peak a - R (g)', 'Passenger Pelvis Peak a - x (g)' 'Passenger Pelvis Peak a - y (g)' 'Passenger Pelvis Peak a - z (g)' 'Passenger Pelvis Peak a - R (g)'};
xlswrite('Results_diadem_dp.xls', header, 'Sheet1', 'A1')
xlswrite('Results_diadem_dp.xls', frontal_info, 'Sheet1', 'A2')
close all