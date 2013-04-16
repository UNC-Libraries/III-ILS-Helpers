# Millennium Payment Information Processor
A helper for working with payment data exported from III Millennium. The Windows .exe version of the script can be run by any Windows user, without the need to install Ruby and other dependencies. 

Takes payment data exported from a Review File of order records. Outputs a tab-delimited .txt file that can be opened with Excel. There are two output types to choose from: 
- *Individual payments* - outputs one line per payment made. Data can be further processed in Excel using PivotTables or other means. Each line contains order record number, fiscal year of payment, (other fields exported), payment data fields.
- *Payments summarized by fiscal year* - outputs one line per order record. Each line contains order record number, (other fields exported), and one column for each fiscal year input by script user. Each of these columns contains total payment amount for that fiscal year.

# Requirements
## Windows .exe version
No need to install Ruby. Use by double-clicking the .exe file. 

The script does require a specific directory set-up for location of the input file, output file, and the script itself. See below in "How to use" section.

## Ruby script (.rb) version
You must have [Ruby] (http://www.ruby-lang.org/en/) installed. This script has been tested on Ruby 1.9.2. Installing Ruby is super-easy; point-and-click .exe installers are [available for Windows] (http://rubyinstaller.org/).

Once Ruby is installed, you will need to install the Ruby Gem called [Highline] (http://highline.rubyforge.org/).

To install this Gem, open the command line shell and type the following commands: 
- gem install highline

# How to use
## Prepare your directory structure
Choose or create a directory/folder in which to place the script (.rb or .exe). This directory can be called whatever you want, but here I'll call it the "ruby_scripts" directory. 

In the ruby_scripts directory, create a new directory called "data". 

In the ruby_scripts directory, create a new directory called "output". 

The structure should look like this: 

```
- ruby_scripts
-- data
-- output
```

Put the payment_info_processor .rb or .exe file(s) in the ruby_scripts directory.

## Prepare your input file
Export from a Review File of order records in Millennium. 

### Required fields
- ORDER - RECORD NUMBER: required, must be first column
- ORDER - PAID: required, must be the last column

You can export whatever fields you want in between RECORD NUMBER and PAID, but only include fields that will be output in a single column.

### Export settings
FIELD DELIMITER:
- Control character (1-127) = 42

TEXT QUALIFIER:
- Text qualifier = None

### Exported file name and location
- File name *must* be payment_data.txt
- Save the file in the "data" folder inside your "ruby_scripts" folder


## Run the script
### .exe version
Just double-click it!

### rb veresion
At command line, from inside ruby_scripts directory: 
- ruby payment_info_processor.rb

## Output
The output files will be put in the "output" directory inside your "ruby_scripts" directory

Open the .txt files with Excel. Choose "65001 : Unicode (UTF-8)" as file origin (encoding) to properly display diacritics.
