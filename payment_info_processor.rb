# coding: utf-8
# Using some standard libraries
require 'date'
require 'pathname'

# Using some external Ruby Gems
require 'rubygems'
require 'highline/import'

# Ocra = One-Click Ruby Application Builder
# This is here for building the Windows executable version
exit if Object.const_defined?(:Ocra)

# There should be a "payment_processor_config.txt" file in the data directory
# It is just 2 lines and looks like:
#   fy_begin_month = 7↩️
#   fy_begin_day = 1↩️
# Not all institutions use the same fiscal year dates, so this makes the script
#  more flexible.
# The weird curved arrow symbols at the end of each line are there to remind you that
#  the computer knows about invisible characters such as line breaks

# Test that config file exists
config_path = Pathname.new("data/payment_processor_config.txt")
unless config_path.exist?
  puts "\n\nCannot find config file: data/payment_processor_config.txt"
  puts "Please see documentation at:"
  puts "   https://github.com/UNC-Libraries/Millennium-Helpers\n\n"
  exit
end

# Load config settings into array
config_lines = []
config_path.each_line do |ln|
  ln.chomp!
  config_lines << ln
  # RESULT:
  # ['fy_begin_month = 7', 'by_begin_day = 1']
end

#make sure the month is set properly
unless config_lines[0].match(/^fy_begin_month = \d\d?\s*$/)
  puts "\n\nFY begin month not properly set in config file."
  puts "Please see documentation at:"
  puts "   https://github.com/UNC-Libraries/Millennium-Helpers\n\n"
  exit
end

#make sure the day is set properly
unless config_lines[1].match(/^fy_begin_day = \d\d?\s*$/)
  puts "\n\nFY begin day not properly set in config file."
  puts "Please see documentation at:"
  puts "   https://github.com/UNC-Libraries/Millennium-Helpers\n\n"
  exit
end

#set the fy start variables
$fystartmonth = config_lines[0].gsub /^.* = /, '' # '7'
$fystartday = config_lines[1].gsub /^.* = /, ''   # '1'

#make sure date created from config month and year is valid
unless Date.valid_date?(2012, $fystartmonth.to_i, $fystartday.to_i)
  puts "\n\nThe month and day in your config file do not combine to create a valid date."
  puts "Please see documentation at:"
  puts "   https://github.com/UNC-Libraries/Millennium-Helpers\n\n"
  exit  
end

def find_fy(adate)
  theyear = adate.year
  fystartnum = Date.new(theyear, $fystartmonth.to_i, $fystartday.to_i).yday
  paydatenum = adate.yday
  if paydatenum >= fystartnum
    fy = theyear.to_i
  else
    fy = theyear.to_i - 1
  end
end

def set_full_year(yr)
  if yr.to_i > 50
    fullyr = '19' + yr
  else
    fullyr = '20' + yr
  end
  return fullyr
end

def get_fy_label(yr)
  thisyear = yr.to_i
  nextyear = thisyear + 1
  label = "FY#{thisyear}-#{nextyear}"
  return label
end

puts "\n\n\n\n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
puts "Welcome to the Millennium Payment Data Processor".upcase
puts "version 1.3.0, 2014-01-23"
puts "written by Kristina Spurgin, ESM, kspurgin@email.unc.edu"
puts "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
puts "\n\nINPUT:"
puts "This program processes payment data exported from a List of"
puts "  order records in Millennium."
puts "\nExported data must include:"
puts "  - order - record number (first field exported)"
puts "  - order - paid (last field exported)"
puts "\nExported data may include:"
puts "  - other fields, exported between record number and paid"
puts "\n\nOUTPUT:"
puts "There are two options for output. Both produce tab-delimited .txt"
puts "  files you can open with Excel. Output options are:"
puts "  - individual payments - outputs one line per payment made. Each"
puts "    line contains order record number, fiscal year of payment,"
puts "    (other fields exported), payment data fields"
puts "  - payments summarized by fiscal year - outputs one line per order"
puts "    record. Each line contains order record number, (other fields"
puts "    exported), and one column for each fiscal year input by script"
puts "    user. Each of these columns contains total payment amount for"
puts "    that fiscal year."
puts "Open the .txt file with Excel. Choose \"65001 : Unicode (UTF-8)\""
puts "    as file origin (encoding) to properly display diacritics." 

puts "\n\n\nDo you want to see instructions on how to export from Millennium"
puts "  in order to use this application? (y/n)"
millinst = gets.chomp
if millinst == "y"

  puts "\n\n\n\n\n\n\n\n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
  puts "HOW TO EXPORT DATA FOR THIS SCRIPT FROM MILLENNIUM"
  puts "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
  puts "REQUIRED FIELDS:"
  puts " -- ORDER - RECORD NUMBER: required, must be first column"
  puts " -- ORDER - PAID: required, must be the last column\n\n"
  puts "Note: you can export whatever fields you want in between."

  puts "\nSet field delimiter:".upcase
  puts " -- Click on \"Field delimiter\". "
  puts " -- In popup dialog, click \"Control character (1-127).\" "
  puts " -- Enter \"42\" (no quotes). "
  puts " -- Click ok. "

  puts "\nSet text qualifier:".upcase
  puts " -- Click on \"Text qualifier\". "
  puts " -- In popup dialog, select \"None.\" "
  puts " -- Click ok. "
  
  puts "\nExport with required name and location so script can find data:".upcase
  puts " -- The exported file MUST be named payment_data.txt"
  puts " -- Save the file in the \"data\" folder inside your"
  puts "    \"rubyscripts\" folder."

  puts "\n\nContinue? (y/n)"
  cont = gets.chomp
  if cont == "n"
    exit
  end
end

puts "\n\n\n\n\n\n\n\n\n\n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
puts "BEFORE CONTINUING, MAKE SURE YOU HAVE:"
puts " - named the exported file \"payment_data.txt\""
puts " - put \"payment_data.txt\" in your rubyscripts\\data folder..."

puts "\n\nAlmost ready to go..."
puts "Do you have any previous results file from this script open now? (y/n)"

results = gets.chomp

if results == "y"
  puts "Please close that file. Is it closed? (y/n)"
  empty = gets.chomp
end

lines = IO.readlines("data/payment_data.txt")

=begin
First two lines of file:
RECORD #(ORDER)*TITLE*FUND*Paid Date*Invoice Date*Invoice Num*Amount Paid*Voucher Num*Copies*Sub From*Sub To*Note
o10002066*Africa research bulletin.*esoci*06-02-10*05-26-10*0106526*1591.57*218304*001*01-01-10*12-31-10*46(01/10)-47(12/10)!B9451699\;02-28-11*02-16-11*0199422*1662.09*224256*001*01-01-11*12-31-11*47(01/11)-48(12/11)!D4685821
=end

lines.each {|l| l.gsub! /"/, '' ; l.chomp!}

#Grab headers, break up meaningfully, and hash for later use
headers = lines.shift.split("*")

=begin
headers = ["RECORD #(ORDER)", "TITLE", "FUND", "Paid Date", "Invoice Date", "Invoice Num", "Amount Paid", "Voucher Num", "Copies", "Sub From", "Sub To", "Note"]

lines = "o10002066*Africa research bulletin.*esoci*06-02-10*05-26-10*0106526*1591.57*218304*001*01-01-10*12-31-10*46(01/10)-47(12/10)!B9451699\;02-28-11*02-16-11*0199422*1662.09*224256*001*01-01-11*12-31-11*47(01/11)-48(12/11)!D4685821"
=end

hdr = {}
hdr[:onum] = headers.shift

=begin
hdr = { :onum = "RECORD #(ORDER)" }
headers = ["TITLE", "FUND", "Paid Date", "Invoice Date", "Invoice Num", "Amount Paid", "Voucher Num", "Copies", "Sub From", "Sub To", "Note"]
=end

hdr[:payment_headers] = headers.pop(9)

=begin
There are NINE payment fields that get repeated on the end of each line
 in the output.

hdr = { :onum = "RECORD #(ORDER)",
        :payment_headers = ["Paid Date", "Invoice Date", "Invoice Num", 
                            "Amount Paid", "Voucher Num", "Copies",
                            "Sub From", "Sub To", "Note" 
                           ]
      }
headers = ["TITLE", "FUND"]
=end

hdr[:other_headers] = headers

=begin
hdr = { :onum = "RECORD #(ORDER)",
        :payment_headers = ["Paid Date", "Invoice Date", "Invoice Num", 
                            "Amount Paid", "Voucher Num", "Copies",
                            "Sub From", "Sub To", "Note" 
                           ],
        :other_headers = ["TITLE", "FUND"]
      }
=end

puts "\n\nWhat output would you like?"

choose do |menu|
  menu.choice :individual_payments do 
    @output_lines = []
    @output_lines << [hdr[:onum], "FY", hdr[:other_headers], hdr[:payment_headers]].flatten.join("\t")
    
=begin
Before we flatten: 
@output_lines = ["RECORD #(ORDER)",
                 "FY",
                 ["TITLE", "FUND"],
                 ["Paid Date", "Invoice Date", "Invoice Num", 
                  "Amount Paid", "Voucher Num", "Copies",
                  "Sub From", "Sub To", "Note" 
                 ]
                ]

After flattening: 
@output_lines = ["RECORD #(ORDER)", "FY", "TITLE", "FUND",
                 "Paid Date", "Invoice Date", "Invoice Num", 
                 "Amount Paid", "Voucher Num", "Copies",
                 "Sub From", "Sub To", "Note" 
                ]

The join turns these into one tab-delimited string: 
"RECORD #(ORDER)\tFY\tTITLE\tFUND\tPaid Date\tInvoice Date\tInvoice Num\tAmount Paid\tVoucher Num\tCopies\tSub From\tSub To\tNote"
=end

=begin
We only have one line in the lines variable, for demo purposes:

lines = "o10002066*Africa research bulletin.*esoci*06-02-10*05-26-10*0106526*1591.57*218304*001*01-01-10*12-31-10*46(01/10)-47(12/10)!B9451699\;02-28-11*02-16-11*0199422*1662.09*224256*001*01-01-11*12-31-11*47(01/11)-48(12/11)!D4685821"
=end
    
    lines.each do |l|
      line = l.split("*")
      
=begin
line = ["o10002066", "Africa research bulletin.", "esoci", "06-02-10", "05-26-10", "0106526", "1591.57", "218304", "001", "01-01-10", "12-31-10", "46(01/10)-47(12/10)!B9451699;02-28-11", "02-16-11", "0199422", "1662.09", "224256", "001", "01-01-11", "12-31-11", "47(01/11)-48(12/11)!D4685821"]
=end

      order_num = line.shift
      
=begin
order_num = "o10002066"

line = ["Africa research bulletin.", "esoci", "06-02-10", "05-26-10", "0106526", "1591.57", "218304", "001", "01-01-10", "12-31-10", "46(01/10)-47(12/10)!B9451699;02-28-11", "02-16-11", "0199422", "1662.09", "224256", "001", "01-01-11", "12-31-11", "47(01/11)-48(12/11)!D4685821"]
=end

      # How many fields come before the payment data starts?
      # However many times that is, shift the next field in the line to :other
      other_data = []
      hdr[:other_headers].size.times do
        other_data << line.shift
      end
      
=begin
order_num = "o10002066"

other_data = ["Africa research bulletin.", "esoci"]

line = ["06-02-10", "05-26-10", "0106526", "1591.57", "218304", "001", "01-01-10", "12-31-10", "46(01/10)-47(12/10)!B9451699;02-28-11", "02-16-11", "0199422", "1662.09", "224256", "001", "01-01-11", "12-31-11", "47(01/11)-48(12/11)!D4685821"]

?? What do you notice that is funky in the data we still have left in the line variable? 
?
?
?
?
?
?
?
?
?
?
?
?
?
=end

      # smoosh payments back together in one string and separate by ;
      payment_lines = line.join("*").split(";")

=begin
line.join('*') produces: 
"06-02-10*05-26-10*0106526*1591.57*218304*001*01-01-10*12-31-10*46(01/10)-47(12/10)!B9451699;02-28-11*02-16-11*0199422*1662.09*224256*001*01-01-11*12-31-11*47(01/11)-48(12/11)!D4685821"

Adding the split on this, we get: 
payment_lines = [
                 "06-02-10*05-26-10*0106526*1591.57*218304*001*01-01-10*12-31-10*46(01/10)-47(12/10)!B9451699",
                 "02-28-11*02-16-11*0199422*1662.09*224256*001*01-01-11*12-31-11*47(01/11)-48(12/11)!D4685821"
                ]
=end
      
      payment_lines.each do |pline|

=begin
The first pline is: 
"06-02-10*05-26-10*0106526*1591.57*218304*001*01-01-10*12-31-10*46(01/10)-47(12/10)!B9451699",
=end
        
        payment = pline.split "*"

=begin
payment = ["06-02-10", "05-26-10", "0106526", "1591.57", "218304", "001", "01-01-10", 
           "12-31-10", "46(01/10)-47(12/10)!B9451699"]
=end

        paid_date = payment[0] # "06-02-10"
        
        pd = paid_date.split "-" # ["06", "02", "10"]

        paidyear = set_full_year(pd[2])
        
=begin
pd[2] = '10'

For reference from above: 
def set_full_year(yr)
  if yr.to_i > 50
    fullyr = '19' + yr
  else
    fullyr = '20' + yr
  end
  return fullyr
end

paidyear = "2010"
=end

        paid_date_f = Date.new(paidyear.to_i, pd[0].to_i, pd[1].to_i)
        # #<Date: 2010-06-02 ((2455350j,0s,0n),+0s,2299161j)>

        fiscalyr = find_fy(paid_date_f)

=begin
For reference from above: 
def find_fy(adate)
  theyear = adate.year
  fystartnum = Date.new(theyear, $fystartmonth.to_i, $fystartday.to_i).yday
  paydatenum = adate.yday
  if paydatenum >= fystartnum
    fy = theyear.to_i
  else
    fy = theyear.to_i - 1
  end
end

theyear = 2010

fystartnum = Date.new(2010, 7, 1).yday
fystartnum = 182 (the 182nd day of the year)

paydatenum = 153

fy (fiscalyr) = 2009
=end
        
        fylabel = get_fy_label(fiscalyr)
        
=begin
For reference from above: 
def get_fy_label(yr)
  thisyear = yr.to_i
  nextyear = thisyear + 1
  label = "FY#{thisyear}-#{nextyear}"
  return label
end

thisyear = 2009
nextyear = 2010
label = "FY#2009-2010"
=end

        @output_lines << [order_num, fylabel, other_data, payment].flatten.join("\t")
      end
      # Done with that one pline (single payment line in an order)
      # Go to the next pline for this order if there is one
    end
    # Done with that order and all its payments

    File.open "output/payments.txt", "wb" do |f|
      @output_lines.each {|l| f.puts l}
    end

    
    puts "\n\n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
    puts "Done!"
    puts "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
    puts "Results are in your \"rubyscripts\\output\" folder."
    puts "Look for a file called \"payments.txt\""
  end

  menu.choice :summary_of_payments_per_fiscal_year do 
    @orders = []
    
    puts "Enter the years for which you want summary data."
    puts "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
    puts "If you want FY2009-10, FY2010-11, and FY2011-12, type:"
    puts "\n\n  2009,2010,2011\n\n"
    puts "Important:\n - the years are separated by commas\n - there are NO spaces in between dates and commas"
    puts " - the years are four digits\n - entering 2009 gets you FY2009-10\n - entering 2010 gets you FY2010-11, and so on."
    puts "\n\nYou can request as many years as you like."
    puts "If there is no data for a year, its values will be zero."
    puts "\n\nType years below and hit enter/return:"

    @@rawyears = gets.chomp.split(",")
    #@@rawyears = ARGV[1].split(",")
    # p rawyears

    class Order
      attr_accessor :onum, :other, :payments, :years
      def initialize(onum)
        @onum = onum
        @other = []
        @payments = []
        @years = {}

        @@rawyears.each do |yr|
          @years[yr] = 0.0
        end
      end

      def calculate
        self.payments.each do |pmt|
          if self.years.has_key? pmt.fy.to_s
            self.years[pmt.fy.to_s] += pmt.amount.to_f
          end
        end
      end
    end

    class Payment
      attr_accessor :pds, :paid_date, :amount, :fy
      def initialize paid_date, amount
        @pds = paid_date
        
        pd = paid_date.split "-"
        pdyr = set_full_year(pd[2])

        @paid_date = Date.new pdyr.to_i, pd[0].to_i, pd[1].to_i
        @amount = amount
        @fy = find_fy(@paid_date)
      end

      # def find_fy date
      #   yr = date.year
      #   mo = date.month
      #   if mo > 6
      #     return yr
      #   else
      #     return yr-1
      #   end
      # end
    end

    lines.each do |l|
      line = l.split("*")
      ord = Order.new(line.shift)
      
      # How many fields come before the payment data starts?
      # However many times that is, shift the next field in the line to :other

      hdr[:other_headers].size.times do
        ord.other << line.shift
      end
      
      # smoosh payments back together in one string and separate by ;
      payment_lines = line.join("*").split(";")
      payments = []
      payment_lines.each do |pline|
        payment = pline.split "*"
        payments << payment
      end

      payments.each do |pmt|
        # 0 element = payment date
        # 3 element = amount
        p = Payment.new pmt[0], pmt[3]
        ord.payments << p
      end
      @orders << ord

    end    

    @orders.each {|ord| ord.calculate}

    # check calculations
    #@orders.each do |ord|
    #  puts "\n\n#{ord.onum}"
    #  puts ord.other[0]
    #  ord.payments.each {|pmt| p pmt.amount if @@rawyears.include? pmt.fy.to_s}
    #  ord.years.each_pair do  |k,v|
    #    yr = k.to_i
    #    nyr = yr + 1
    #    puts "FY#{yr.to_s}-#{nyr.to_s}: $#{v}"
    #  end
    #end


    output = []
    yrlabels = []
    @@rawyears.each do |yr|
      yrlabels << get_fy_label(yr)
    end
    output << [hdr[:onum], hdr[:other_headers], yrlabels].flatten.join("\t")
    @orders.each do |ord|
      out = []
      out << ord.onum
      out << ord.other
      ord.years.each_value {|yr| out << yr}
      output << out.flatten.join("\t")
    end

    File.open "output/payment_summary.txt", "wb" do |f|
      output.each {|l| f.puts l}
    end

    puts "\n\n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
    puts "Done!"
    puts "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
    puts "Results are in your \"rubyscripts\\output\" folder."
    puts "Look for a file called \"payment_summary.txt\""
  end
end

puts "Press return/enter to exit."
done = gets
exit
