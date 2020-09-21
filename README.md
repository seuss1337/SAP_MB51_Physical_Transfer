# SAP physical document generator

## Script to automate physical record requirements for material transfer within my company.

For auditing reasons we are required to keep physical copies of material transfers done through
SAP, because material movements for profit recognition, cost assignment, statement balancing 
and date errors, storage locations can have a large number of transfers. Since these documents
are created manually, this becomes a tedious and time consuming process.  This script
reads an excel output export of SAP MB51 pulls all material transfers and fills out an excel 
transfer sheet for all material transfer documents and saves them to an excel file. One then
only needs to print out the excel file as opposed to filling each document by hand.


# still under development, currently only works for 1 storage location at a time.
