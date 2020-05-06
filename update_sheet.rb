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


# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

config = YAML.load(File.read("config.yml"))       

puts "Loaded spreadsheet_id #{config[0]['spreadsheet_id']}"

#Add rows to spreadsheet


months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
          'Agosto','Septiembre','Octubre', 'Noviembre', 'Diciembre']

day_num = Time.now.strftime("%d").to_i
month_num = Time.now.strftime("%m").to_i


puts "update for #{months[month_num-1]} #{day_num}"


range_fmt = "#{months[month_num-1]}!B#{day_num}:F#{day_num}"
range_name = [range_fmt]


values = [
  [day_num, 'ahora','valor','pepe','valor']
]


values_range = Google::Apis::SheetsV4::ValueRange.new(values: values)


response = service.update_spreadsheet_value(config[0]['spreadsheet_id'], range_name, values_range, value_input_option: 'USER_ENTERED')
