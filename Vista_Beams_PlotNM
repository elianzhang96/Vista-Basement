from openpyxl import load_workbook, workbook
from gsapy import GSA
from datetime import datetime
import Lib_Beams_PlotNM as lib
import os
import subprocess
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side, PatternFill

# GSA model and outputs
gsa_file = r'C:\Users\Elian.Zhang\Arup\PROJECT VISTA - 4 Internal Data\04 Calculations\07 Structural\09 Substructure\Stage 4\12 Basement Design\01_GSA\03_SIMPLIFIED\Vista_BasementSlab_Simplified_GL_v14.gwb'

assembly_list = [75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127]
#assembly_list = [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127]
#case_list = ["C101"]
case_list = ["C218", "C219", "C220", "C221", "C226", "C227", "C228", "C229", "C318", "C319", "C320", "C321", "C326", "C327", "C328", "C329"]


######################################################################################################

start_time = datetime.now() #Start timer

# Create output folder name
formatted_datetime = start_time.replace(second=0).strftime("%Y-%m-%d %H-%M")
outputFolder_excel = rf'C:\Temp\{formatted_datetime} Vista_beam NM plots'
outputFileName_excel = f'Vista_BeamNM_Results {formatted_datetime}.xlsx'

################# EXTRACT RESULTS FROM GSA

print("Opening GSA file. Time: ", datetime.now().replace(microsecond=0))
model = GSA(gsa_file, version="10.2")
sheetname = 'GSA_assembly_'
gsa_model_name = gsa_file.split("\\")[-1]

assembly_results_dict = lib.extract_from_gsa(assembly_list, case_list, model) # Create a dictionary where to store the GSA results
print("Results extracted from GSA. Time: ", datetime.now().replace(microsecond=0))

if not os.path.exists(outputFolder_excel): # Create the output folder
   os.makedirs(outputFolder_excel)


################# FIND OUTLIER VALUES TO CHECK THE BEAM SECTION AGAINST

# Create a dict of sub beams (thirds of beams) and assign the first third of results to the first sub beam, the second third to the second sub beam,etc.)
sub_beams = ["end1", "mid", "end2"]
assembly_results_dict = lib.split_results_subbeams(assembly_results_dict, sub_beams)

# Find outlier results (critical results) for each sub beam using convex hull
plot_folder = outputFolder_excel+"\\NM plots"
if not os.path.exists(plot_folder): # Create the output folder for NM plots
    os.makedirs(plot_folder)

sub_assembly_crit_results_dict = lib.find_outliers(assembly_results_dict, sub_beams, plot_folder)

################# WRITE RESULTS IN EXCEL

# Define a style with one decimal place for numbers
number_style = NamedStyle(name="one_decimal", number_format='0.0')

# Critical results
lib.write_criticalresults_to_excel(outputFileName_excel, "NM - Outlier Assembly Forces", gsa_model_name, sub_assembly_crit_results_dict, outputFolder_excel, number_style, start_time)

# All assembly results
lib.write_results_to_excel(sheetname, gsa_model_name, assembly_results_dict, outputFolder_excel, outputFileName_excel, number_style, start_time) # Write the GSA result dictionary in an Excel spreadsheet to read the GSA results more easily (optional)



################# OPEN FOLDER

end_time = datetime.now()
delta=end_time-start_time
print("Elapsed time: ", str(delta).split(".")[0])

subprocess.Popen(['explorer', outputFolder_excel], shell=True)

################# CLOSE GSA FILE

del model