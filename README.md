# ml_metrics_to_google_spreadsheet

Push mercadolibre reputation values to a google spreadsheet

## intro 


- `cretate_sheet.rb` runs once a year, i will create a new spreadsheet with the current year prefixed and also creates one sheet for every month with columns. 

- `update_sheet` runs everyday, it fetchs the daily reputations metrics of the user and push the values to the spreadsheet on the corresponding sheet (month) and the row based on the day of the month

## installation 

As usual, clone and `bundle install` 

To fix a deps issue this scrips uses a custom-compiled version of the mercadopago gem
with faraday and json deps updated to current versions. 


## config 

config is a little bit tricket, you will have to grant access to your app from mercadolibre and also from google. 

### mercadolibre 

1. Create a `config.yml` file on the working dir of the script and fill the ML authentication data. Take a look at config.sample.yml for reference 


### google drive/spreadsheets












  
