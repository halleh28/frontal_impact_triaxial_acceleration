load('IIHS_plots','frontal_info')

header = {'Test ID' 'Vehicle Info' 'Vehicle dV (kph)' 'Vehicle Peak a (g)' 'Head Peak a - x (g)' 'Head Peak a - y (g)' 'Head Peak a - z (g)' 'Head Peak a - R (g)', 'Chest Peak a - x (g)' 'Chest Peak a - y (g)' 'Chest Peak a - z (g)' 'Chest Peak a - R (g)'};
xlswrite('Results_plots.xls', header, 'Sheet1', 'A1')
xlswrite('Results_plots.xls', frontal_info, 'Sheet1', 'A2')
close all