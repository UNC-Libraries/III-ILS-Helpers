**Quick download links:**
- [Windows 10 .exe](http://infomuse.net/millennium-helpers/payment_info_processor_win10.exe)
- [pre-Windows 10 .exe](http://infomuse.net/millennium-helpers/payment_info_processor.exe)
- [Ruby script (.rb) .zip](https://github.com/UNC-Libraries/Millennium-Helpers/archive/master.zip)

**Script not working for you?** Check [here](https://github.com/UNC-Libraries/III-ILS-Helpers#frequent-issues-that-prevent-script-from-running-properly)

# Millennium/Sierra Payment Information Processor
A helper for working with payment data exported from III Millennium or Sierra. The Windows .exe version of the script can be run by any Windows user, without the need to install Ruby and other dependencies.

Takes payment data exported from a Review File of order records. Outputs a tab-delimited .txt file that can be opened with Excel. There are two output types to choose from:
- *Individual payments* - outputs one line per payment made. Data can be further processed in Excel using PivotTables or other means. Each line contains order record number, fiscal year of payment, (other fields exported), payment data fields.
- *Payments summarized by fiscal year* - outputs one line per order record. Each line contains order record number, (other fields exported), and one column for each fiscal year input by script user. Each of these columns contains total payment amount for that fiscal year.

# Set up required before first use
## ALL versions (Windows .exe and Ruby script (.rb))
### Prepare your directory structure
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

<i>Note for more advanced users: You can run the script from anywhere; there doesn't have to be a ruby_scripts directory. BUT, there must be a data directory and an output directory in whatever directory you are running the script from. It's clunky and inflexible, but it is the only way I know to make the .exe version work for my colleagues who don't know how to use the command line.</i>

### Prepare your configuration file (First use only)
The configuration file will tell the script when your new fiscal year begins. I will use my institution as an example. Our fiscal year runs from July 1 to June 30.

In the ruby_scripts/data directory, create a new text file named payment_processor_config.txt

payment_processor_config.txt will consist of two lines. Except for the number at the end of the line, your text should match what is below exactly. Copy/paste it in to be sure.

```
fy_begin_month = 7
fy_begin_day = 1
```

Change the number at the end of each line to reflect when your new fiscal year begins. Save and close payment_processor_config.txt.

## Windows .exe version
Download Windows .exe file:
- [pre-Windows 10](http://infomuse.net/millennium-helpers/payment_info_processor.exe)
- [Windows 10](http://infomuse.net/millennium-helpers/payment_info_processor_win10.exe)

If a window pops up and asks you whether to save or open the file, select save.

Save the file into your ruby_scripts directory.

## Ruby script (.rb) version
- Install [Ruby](http://www.ruby-lang.org/en/). Point-and-click Windows .exe Ruby installers are [available for Windows](http://rubyinstaller.org/).

- Install the Ruby Gem called [Highline] (http://highline.rubyforge.org/). To install this Gem, open the command line shell and type the following commands:

```
gem install highline
```

- Download the script
 - Click the "Download ZIP" button in the righthand column of https://github.com/UNC-Libraries/Millennium-Helpers
 - Save file to your computer
 - Unzip the file you just saved
 - Move payment_info_processor.rb to ruby_scripts directory


# Regular use
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

REPEATED FIELD DELIMITER (leave as default):
- ASCII character = ;

### Exported file name and location
- File name *must* be payment_data.txt
- Save the file in the "data" folder inside your "ruby_scripts" folder


## Run the script
### .exe version
Just double-click it!

### rb version
At command line, from inside ruby_scripts directory:
- ruby payment_info_processor.rb

## Output
The output files will be put in the "output" directory inside your "ruby_scripts" directory

Open the .txt files with Excel. Choose "65001 : Unicode (UTF-8)" as file origin (encoding) to properly display diacritics.


# Frequent issues that prevent script from running properly
If the script is failing to work for you, check the following:

- Did you export your payment_data.txt file from a list of bib records (or other non-order record types)? Unfortunately you have to do the export from a list of order records.
- Is there a typo in your input file name? It must be: payment_data.txt
- And, if you just got a new computer:
  - Did you [define the beginning of your library's fiscal year in the configuration file](https://github.com/UNC-Libraries/III-ILS-Helpers#prepare-your-configuration-file-first-use-only)?
  - If new computer is running Windows 10, did you get the [Windows 10 version of the script](http://infomuse.net/millennium-helpers/payment_info_processor_win10.exe)?

If none of those are the culprit, please go [here](https://github.com/UNC-Libraries/III-ILS-Helpers/issues) and open a new issue. Please attach the following:
- your payment_data.txt input file
- a screenshot of your Create Lists export settings before you start the export

Please summarize the script options you are using---individual payments vs. summary by fiscal year. If the latter, please say which fiscal years you input.
