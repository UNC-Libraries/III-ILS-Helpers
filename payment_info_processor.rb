require 'rubygems'
require 'highline/import'
require 'date'
require 'pathname'

exit if Object.const_defined?(:Ocra)

#test that config file exists
config_path = Pathname.new("data/payment_processor_config.txt")
unless config_path.exist?
  puts "\n\nCannot find config file: data/payment_processor_config.txt"
  puts "Please see documentation at:"
  puts "   https://github.com/UNC-Libraries/Millennium-Helpers\n\n"
  exit
end

#load config settings into array
config_lines = []
config_path.each_line do |ln|
  ln.chomp!
  config_lines << ln
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
$fystartmonth = config_lines[0].gsub /^.* = /, ''
$fystartday = config_lines[1].gsub /^.* = /, ''

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
lines.each {|l| l.gsub! /"/, '' ; l.chomp!}

#Grab headers, break up meaningfully, and hash for later use
headers = lines.shift.split("*")
# p headers
hdr = {}
hdr[:onum] = headers.shift
# 9 is the number of fields of payment data that get repeated.
hdr[:payment_headers] = headers.pop(9)
hdr[:other_headers] = headers
#p hdr[:payment_headers]
#p hdr[:other_headers]
#p hdr

puts "\n\nWhat output would you like?"

choose do |menu|
  menu.choice :individual_payments do 
    @output_lines = []
    @output_lines << [hdr[:onum], "FY", hdr[:other_headers], hdr[:payment_headers]].flatten.join("\t")
    
    lines.each do |l|
      line = l.split("*")
      order_num = line.shift
 
      # How many fields come before the payment data starts?
      # However many times that is, shift the next field in the line to :other
      other_data = []
      hdr[:other_headers].size.times do
        other_data << line.shift
      end
 
      # smoosh payments back together in one string and separate by ;
      payment_lines = line.join("*").split(";")
      payment_lines.each do |pline|
        payment = pline.split "*"

        paid_date = payment[0]
        pd = paid_date.split "-"

        paidyear = set_full_year(pd[2])
        
        paid_date_f = Date.new paidyear.to_i, pd[0].to_i, pd[1].to_i

        fiscalyr = find_fy(paid_date_f)
        fylabel = get_fy_label(fiscalyr)
        
        @output_lines << [order_num, fylabel, other_data, payment].flatten.join("\t")
      end
   end

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
