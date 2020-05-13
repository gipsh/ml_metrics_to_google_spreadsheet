require "google/apis/drive_v3"
require "googleauth"
require "googleauth/stores/file_token_store"
require "fileutils"
require 'google/apis/sheets_v4'
require 'yaml'


OOB_URI = "urn:ietf:wg:oauth:2.0:oob".freeze
APPLICATION_NAME = "ML metrics to spreadsheet".freeze
CREDENTIALS_PATH = "credentials.json".freeze

# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = "token.yaml".freeze
#SCOPE = Google::Apis::DriveV3::AUTH_DRIVE_METADATA_READONLY

SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS
##
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  client_id = Google::Auth::ClientId.from_file CREDENTIALS_PATH
  token_store = Google::Auth::Stores::FileTokenStore.new file: TOKEN_PATH
  authorizer = Google::Auth::UserAuthorizer.new client_id, SCOPE, token_store
  user_id = "default"
  credentials = authorizer.get_credentials user_id
  if credentials.nil?
    url = authorizer.get_authorization_url base_url: OOB_URI
    puts "Open the following URL in the browser and enter the " \
         "resulting code after authorization:\n" + url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
  end
  credentials
end

def create_month_sheet(service, spreadsheet_id, month_name) 
  puts " - Creating sheet for #{month_name}"
  column_count = 55

  add_sheet_request = Google::Apis::SheetsV4::AddSheetRequest.new
  add_sheet_request.properties = Google::Apis::SheetsV4::SheetProperties.new
  add_sheet_request.properties.title = month_name

  grid_properties = Google::Apis::SheetsV4::GridProperties.new
  grid_properties.column_count = column_count
  add_sheet_request.properties.grid_properties = grid_properties

  batch_update_spreadsheet_request = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new
  batch_update_spreadsheet_request.requests = Google::Apis::SheetsV4::Request.new

  batch_update_spreadsheet_request_object = [ add_sheet: add_sheet_request ]
  batch_update_spreadsheet_request.requests = batch_update_spreadsheet_request_object
  response = service.batch_update_spreadsheet(spreadsheet_id,
  batch_update_spreadsheet_request)
 
  # Create cols
  cols = [ [ 'DIA','DESPACHO', 'DEMORAS','CANCELADAS', 'RECLAMOS EN MEDIACIÓN','CALIDAD DE ATENCIÓN'] ]
  range_fmt = "#{month_name}!B2:G2"
  range_name = [range_fmt]
  col_range = Google::Apis::SheetsV4::ValueRange.new(values: cols)

  result = service.update_spreadsheet_value(spreadsheet_id,
                                          range_name,
                                          col_range,
                                          value_input_option:'RAW')



end

# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

request_body = Google::Apis::SheetsV4::Spreadsheet.new()
props =  Google::Apis::SheetsV4::SpreadsheetProperties.new()
props.title = "#{Time.now.strftime("%Y")} ML puntuacion"
request_body.properties = props

response = service.create_spreadsheet(request_body)
spreadsheet_id = response.spreadsheet_id

puts "title: #{request_body.properties.title}"
puts "id: #{spreadsheet_id}"


# Add months

months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 
	  'Agosto','Septiembre','Octubre', 'Noviembre', 'Diciembre']


months.each { |m| create_month_sheet(service, spreadsheet_id,m) }


puts "saving spreadsheet_id to config.yaml"
config = [ "spreadsheet_id" => spreadsheet_id, "name" => request_body.properties.title ]
File.open("config.yml", "w") { |file| file.write(config.to_yaml) }
