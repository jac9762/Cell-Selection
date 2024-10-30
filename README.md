# Cell-Selection

This repository contains two scripts designed to be used chronologically. The function is to expedite the manual sorting and analysis of regions of interests of single cells within multi-channeled images.

Both corresponding scripts, 'CELL_SELECTION.CPPIPE' and 'PROCESS_IMAGES.PY', can be found below or in their respective folder.




__________________________________________________




CELL_SELECTION.CPPIPE v1.0

DR. ALDEBERAN HOFER LABORATORY

PUBLISHED OCT 31, 2024


SCRIPT DESCRIPTION

This script is designed to take 1+ groupings of 4 two-dimensional images from multiple channels of the same image.

Next, ROIs are automatically drawn for cell nuclei, and channel data is standardized for user visualization (only raw image data used for analyses).

Finally, the script requests the users to view and manually sort any ROIs into one of four output groups: Ciliated Non-Expressing Cells, Non-Ciliated Non-Expressing Cells, Ciliated Expressing Cells, Non-Ciliated Expressing Cells.




SOFTWARE DOWNLOAD

This script was written using Cell Profiler v4.2.6. Download the same or newer version of Cell Profiler at https://cellprofiler.org/releases before continuing. Cell Profiler can run on macOS, and Windows.

Note: Upgrading your computers CPU and RAM will increase the speed of analysis and number of images input at once.





CONTINUE AFTER CELL PROFILER DOWNLOAD

Open Cell Profiler Application. There will be a blank area labeled ‘Drop a pipeline here (.cppipe or .cpproj) or double-click to add modules ’ on the left of the application window. Drag the file ‘Cell_Selection.cppipe’ into the blank area. Cell Profiler will auto-fill the modules to be used for this script.





INPUT IMAGES

Input images files must be 4 two-dimensional .tif files per experiment.

Input image file names are REQUIRED to start with ‘C1’ ‘C2’ ‘C3’ ‘C4’ according to their respective channel. These channels should be consistent, but can be adjusted each run.

The rest of the image file names should be identifiable to their corresponding experiment. Examples: ‘C1_Experiment2.tif’ and ‘C4_Experiment2.tif’ would be grouped together.

Naming the files in this manner is REQUIRED, because Cell Profiler needs to group images together as one experiment, and differentiate between Channels 1-4 within each experiment.

Once files are properly named, they may be dragged and dropped into the blank area labeled ‘Drop files and folders here’ on the top of the application window. Make sure ‘Images’ is selected in the top-left of application window so the blank space for input images is available.

Note: It is very important the file naming system is correct. Double check if uncertain.





SCRIPT SETTINGS

There are three primary settings which may need to be adjusted by the user. On the right side of the Cell Profiler application window, there are six modules with an open eye next to them filled in, four are labeled ‘ImageMath’ and the other two are labeled ‘GrayToColor’ and ‘ExportToSpreadsheet’ respectively.

Click each ‘ImageMath’ module to alter/confirm which channel numbers correspond with the four stains: DNA, Primary Cilia, Transfected Construct, and Histone Marker. These selections should match with the images ‘C1’ ‘C2’ ‘C3’ ‘C4’ respectively. It is important these channels are correct or else the resulting data will be incorrectly sorted or calculated.

Click the ‘GrayToColor’’ module to adjust the relative weight of each color for the user manual selection. The user may need to adjust these values between experiment to adjust for relative levels of fluorescence. (Blue – DNA; Red – Construct; Green – Cilia)

Click the ‘ExportToSpreadsheet’’ module to change the path for the exported data to be saved. This is labeled ‘Output file location’ and can be found at the top of the module after clicking ‘ExportToSpreadsheet.’

Note: Make sure the path for your exported data is a valid pathname on your device.





RUN SCRIPT

Double check input images and pipeline have been added to the application window. Double check all images are named correctly. Once the user is ready, press the button ‘Analyze Images’ in the bottom left of the application window.

An estimated total time and percent completion will be displayed in the bottom-right corner of the Cell Profiler application window for the user to stay updated on progress.

The data will be output in multiple .csv files into the export path defined by the user. This data is ready to be used by the next script ‘Process_Data.py’ to be consolidated and saved as an .xlsx file.


__________________________________________________



PROCESS_IMAGES.PY v1.0

DR. ALDEBERAN HOFER LABORATORY

PUBLISHED JULY 24, 2024




SCRIPT DESCRIPTION

This script is design to receive the output data from ‘Cell_Selection.cppipe’, extract the relevant data and present it in a user-friendly format with data sorted and exported into a single Excel document.




SOFTWARE DOWNLOAD

This script was written in the Python v3.10.9 using Thonny IDE v4.0.2. Make sure Python is downloaded at https://www.python.org/downloads. Any Python IDE can be used to access the script. Thonny is a free open source Python IDE that can be downloaded at https://thonny.org. 





SCRIPT SETTINGS

Open script with your Python IDE (Thonny). The very first two lines of the program will read:

1) input_path = "/Users/Desktop/Automated_Histone_Modification/Data/Excel_Files/CSV/Data/"

2) output_path = "/Users/Desktop/Automated_Histone_Modification/Data/Excel_Files/XLSX/Processed/"

Copy the path of the output data folder from the previous script (Cell_Selection.cppipe). Paste this path to replace the path on the first line. 

Locate your desired output data folder, or create your own. Copy the pathname and paste this path to replace the path on the second line.




RUN SCRIPT

Open script with your Python IDE (Thonny). Double check the input and output file paths are the desired paths.

Run the script ‘Process_Data.py’ on your Python IDE. “DATA PROCESSING DONE” will be printed in the once the script has been completed.

Check your output data folder for the finalized Excel output data, named ‘Processed_Histone_Ratio_Data.xlsx’.
