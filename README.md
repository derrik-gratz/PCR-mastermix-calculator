# PCR_Reagent_Mastermix_calculator
Pulls PCR assay information and samplecounts from a daily generated list and calculates mastermix quantities based on the day's demand

This project is for an animal genetics laboratory. Our LIMS outputs all of the disease tests needed and every animal that ordered 
that test. This information is modeled in the "Platemap" file. This program was designed to automate the PCR setup for all of these samples.
The program counts all of the samples associated with each assay. Note that the second worksheet of the platemap does list samples below
each assay/primer pair, but we cannot use that information because after LIMS spits out an initial worklist, we add/remove many 
samples and tests depending on many circumstances. 
A separate file pairs all of the disease tests, which each have their own PCR primer set, with a specific PCR enzyme 
mastermix (only two as of now, Qiagen and Zymo). 
Another file is a template for the output, with excel formulas to do the math for mastermix recipees based on the number of samples.
