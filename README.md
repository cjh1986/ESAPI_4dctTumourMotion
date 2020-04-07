# ESAPI_4dctTumourMotion
ESAPI C# code to determine approximate tumour motion from a 4DCT scan.

This code was created as a tool for gathering data on pre-treatment 4DCT scans for radiotherapy to assess the typical work load for a department. It was to see how often tumours have large motion and so how often gated treatments are typically required per year and to inform staff about which tumour sizes and locations move the most.  

Code opens patients from a list of IDs and runs through course and plans to find images that fit certain criteria e.g from a clinical plan, used for treatment delivery, not a QA plan amongst other things. The code then determines the centre of mass for each GTV delineated on phase images and determines the relative motion in each axis and returns this. The code also returns the poisition of the GTV relative to the centre of mass of full lung volume to make an estimate of tumour location. This is output as a csv file.
