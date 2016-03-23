#!/usr/bin/env ruby

require 'axlsx'
require 'optparse'
require 'ostruct'
require 'Rmagick'
require 'progress_bar'
include Magick


options = {}
optparse = OptionParser.new do |opts|
	opts.banner = "Usage: imgtoxlsx.rb [options]"
	opts.banner += "\nExample: imgtoxlsx -i a.jpg,b.jbg -o c.xlsx"
	opts.separator ""
	opts.separator "Required options:"

	options[:in] = nil
	opts.on("-i", "--input [IMAGE]", Array, "comma separated list of images") do |input|
		options[:in] = input
	end

	opts.separator "Options:"
	options[:out] = "out.xlsx"
	opts.on("-o", "--output [OUTPUT_FILE]",
		"output file name to use (default: #{options[:out]}") do |output|
		options[:out] = output
	end



	opts.on('-h', '--help', 'Display this screen') do
		puts "\n\n#{optparse}\n\n"
		exit
	end
end


optparse.parse!
fuck_ms_and_their_ratio = 5.64
height = 3
width = height / fuck_ms_and_their_ratio

unless options[:in]
	puts "\n\n#{optparse}\n\n"
	exit
end

def value_to_16_bit_hex(value)
	value = value.to_s(16)
	value = "0" * (4 - value.length) + value unless value.length > 3
	value[0,2].upcase
end
def get_row_styles(sheet, old_styles, image, row_number)
	styles = []
	image.columns.times do |i|
		pixel = image.get_pixels(i, row_number, 1, 1)[0]
		pixel = "FF#{[pixel.red, pixel.green, pixel.blue].collect{|v| value_to_16_bit_hex(v)}.join("")}"
		old_styles[pixel] ||= sheet.styles.add_style(:bg_color => pixel)
		styles << old_styles[pixel]
	end
	styles
end

Axlsx::Package.new do |file|

	options[:in].each do |file_name|
		file.workbook.add_worksheet(:name => file_name) do |sheet|
			styles = {}
			puts "processing #{file_name}"
			image = ImageList.new(file_name)[0]
			bar = ProgressBar.new(image.rows)
			image.rows.times do |num|
				sheet.add_row ([nil] * image.columns), 
					:style => get_row_styles(sheet, styles, image, num),
					:height => height
				bar.increment!
			end
			sheet.column_widths *([width]*image.columns)
			sheet.sheet_view.show_grid_lines = false
		end
	end
	puts "Saving to #{options[:out]} (can't generate progress bar long enough)"
	file.serialize(options[:out])
end


