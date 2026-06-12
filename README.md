# Raisebore-Rig-Planning-App
This is a windows app to help the planner to see how incoming Tenders/Jobs fit existing assets' availability.
The data source and algorithm chartflow are attached as image within this repository for visualization purpose.

There are 3 main availability aspects being tested here named dates, reamer size capability, and rod availability.
This specific algorithm will runthrough each new job being tested and prioritized by the data-entry sequence.
It also will runthrough each rig/equipment started with Rig A, Rig B, etc until there is a match.

For "Dates", each raisebore rig/equipment has projects with dates assigned to it. Name Rig A, Rig B, Rig C, etc.
Source of the data comes from Scheduling File.

For "Reamer Size Capability", each rig has certain reamer size(s) that it can handle with corresponding rod size(s).
Source of the data comes from Rig Capability File comprising rig names with certain reamer size(s) it can accommodate.
There also should be an information of certain rod size(s) that can support certain reamer size(s). 
It is a many-to-many relationship here.

For "Rod Availability", each job will come with target meters, which this app will go through the inventory file 
looking for available supported rod(s) and rod type(s). The inventory file tells you the count of available and
unavailable rods to use. Combined with the length for each rod type, you can compare the required meters/length 
for the job.

So, you will roughly need 3 files as input comprising Rig Schedule, Rig Capability, and Rod Inventory which
I will attach the excel file for illustration purpose.
