# BGS_boreholelogs_extractor
## Purpose:
The aim of the script is to automate the process of extracting historical borehole logs from the [BGS website](https://www.bgs.ac.uk/) within a defined location. This should eliminate the process of having required to manually extract the logs individually through the website.

## Whats in the file:
1. BGS_boreholelogs_extractor.exe - Executable file (Packaged python script)
2. BGS_boreholelogs_extractor.py - Python script
3. BGS_boreholelogs_extractor.spec - spec file for generating the Excutable file using Pyinstaller
4. GeoIndexData.txt - Sample data file that can be extracted from BGS using [GeoIndex](https://www.bgs.ac.uk/data/mapViewers/home.html?src=topNav)

## How to use (Using the Excutable file - recommended):
1. Navigate to [GeoIndex](http://mapapps2.bgs.ac.uk/geoindex/home.html) and click **Add Data**.
2. Under the **Borehole Section**, click **Addition Sign** for Borehole Scans.
3. Navigate to site of interest and **Enlarge Site of Interest** until you are able to see the boreholes marked clearly on the map. See Figure below.
![Figure 1](/images/image1.png)
4. Under the **Search Tab**, click **Polygon Spatial** query button and **Draw** the area of interest on the map. See Figure below.
![Figure 2](/images/image2.png)
5. Click **Export** selection which will download the "GeoIndexData.txt" file. The text file contains the list of boreholes located within the area of interest.
6. **Save** the **"GeoIndexData.txt"** file together **with** the **"BGS_boreholelogs_extractor.exe"**.
7. **Run** the **BGS_boreolelogs_extractor.exe** file.
8. Upon completion of the downloading process, you should see following **Outputs** in the folder:
   - Borehole_logs folder - Contains all the borehole scans downloaded from the website
   - Borehole_logs.docx - A word document with compilation of all the downloaded logs
   - Data - Excel version of the "GeoIndexData.txt" file
   - error_log - text file that records which borehole scans that haven't been downloaded due to: 1) The borehole scan needs to be purchased from the [BGS website](https://www.bgs.ac.uk/); or 2) 403 error from the website (normally due to the fact that BGS doesn't allow file that is > 0.5mb to be downloaded)

## How to use (Using the Python script):
1. If you do not have Python installed on your system, download python via [Anaconda](https://www.anaconda.com/distribution/).
2. Install Python-docx module. Follow the steps outlined [here](https://python-docx.readthedocs.io/en/latest/user/install.html).
3. Follow Step 1 - 5 from the above section.
4. **Save** the **"GeoIndexData.txt"** file together **with** the python script **"BGS_boreholelogs_extractor.py"**.
5. Open command prompt and navigate to the folder where you are working on via the command prompt. (Use cd to navigate)
6. On the command prompt type "python BGS_boreholelogs_extractor.py" and hit enter.
7. Upon completion of the downloading process, you should see following **Outputs** in the folder:
   - Borehole_logs folder - Contains all the borehole scans downloaded from the website
   - Borehole_logs.docx - A word document with compilation of all the downloaded logs
   - Data - Excel version of the "GeoIndexData.txt" file
   - error_log - text file that records which borehole scans that haven't been downloaded due to: 1) The borehole scan needs to be purchased from the [BGS website](https://www.bgs.ac.uk/); or 2) 403 error from the website (normally due to the fact that BGS doesn't allow file that is >0.5mb to be downloaded)

