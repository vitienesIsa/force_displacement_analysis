# Mechanical testing force-displacement curve analysis

Developed by Isabela Vitienes as part of the publication Vitienes I.*, Ross E.* et al., Bone, 2026

Iterates through per-measurement CSV data, plots each force-displacement curve, and prompts user to define (by clicking in plot) the following points:
   1. x_0 - start of loading
   2. x_yield - yield point, i.e. transition from elastic to plastic deformation
   3. x_ult - ultimate load, i.e. peak load
   4. x_fail - fracture/failure, i.e. when load drops to/near zero

X-coordinates are recorded and correspoinding y-values are extracted from csv.

Results are updated and saved to an excel file after each plot. Therefore, progress is not lost if script is interrupted. Already-processed measurements are skipped. 

 INPUT: root path to folder containing measurements
 OUTPUT: .xlsx sheet stored in root path with x- and y-coordinates of selected points, and derived outcome measures:
            - Yield force (N)
            - Ultimate force (N)
            - Failure force (N)
            - Stiffness (N/mm) 
            - Post-yield displacement (mm)
            - Work-to-failure (N*mm)
            - Sampling rate (Hz)
            - Measurement duration (s)
** Note: 
   Script assumes that data is stored in root path as follows:
   
   Root path 
       |
       ---Folder, date of measurement
               |
               ---Folder, specimen-specific measurements
                       |
                       --- .csv file
