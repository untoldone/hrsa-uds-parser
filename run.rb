# Input file from https://www.hrsa.gov/foia/electronic-reading.html section 2019 UDS Data

require "roo"

if ARGV[0] == nil
  puts "First argument must be UDS dataset file path"
end

if ARGV[1] == nil
  puts "Second argument must be output file path"
end


xlsx = Roo::Excelx.new(ARGV[0])


def build_bhcmis_index(xlsx, sheet_name, header_row_coordinate, data_start_row_coordinate)
  index = {}
  sheet = xlsx.sheet(sheet_name)
  id_index = sheet.row(header_row_coordinate).find_index("BHCMIS ID")
  sheet.each_row_streaming(offset: data_start_row_coordinate - 1) do |row|
    index[row[id_index].value] = row[id_index].coordinate[0]
  end

  index
end

def fetch_row(xlsx, sheet_name, header_row_coordinate, row_coordinate)
  sheet = xlsx.sheet(sheet_name)
  headers = sheet.row(header_row_coordinate)
  values = sheet.row(row_coordinate)

  result = {}
  headers.each_with_index do |header, index|
    result[header] = values[index]
  end

  result
end

input_indexes = {}
output_headers = xlsx.sheet("HealthCenterSiteInfo").row(1)

i = build_bhcmis_index(xlsx, "HealthCenterInfo", 1, 2)
input_indexes["HealthCenterInfo"] = i
output_headers.concat(xlsx.sheet("HealthCenterInfo").row(1))

data_sheet_names = ["Table3A",
"Table3B",
"Table4",
"Table5",
"Table6A",
"Table6B",
"Table7_1",
"Table7_2",
"Table8A",
"Table9D",
"Table9E",
"HITInformation",
"OtherDataElements",
"Workforce"]

data_sheet_names.each do |sheet_name|
  i = build_bhcmis_index(xlsx, sheet_name, 1, 3)
  input_indexes[sheet_name] = i
  output_headers.concat(xlsx.sheet(sheet_name).row(1))
end

output_headers.uniq!

site_sheet = xlsx.sheet("HealthCenterSiteInfo")
row_count = site_sheet.last_row

CSV.open(ARGV[1], "w:iso-8859-1") do |output|
  output << output_headers

  (row_count - 1).times do |index|
    data = fetch_row(xlsx, "HealthCenterSiteInfo", 1, index + 2)
    bhcmis_id = data["BHCMIS ID"]
    row = Array.new(output_headers.length)

    input_indexes.each do |sheet_name, row_indexes|
      if row_indexes[bhcmis_id]
        row_data = fetch_row(xlsx, sheet_name, 1, row_indexes[bhcmis_id])

        data = data.merge(row_data)
      end
    end

    output_headers.each_with_index do |header, index|
      row[index] = data[header].force_encoding("iso-8859-1")
    end

    output << row
  end
end
