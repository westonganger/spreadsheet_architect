Dummy::Application.routes.draw do
  get 'spreadsheets/csv', to: 'spreadsheets#csv'
  get 'spreadsheets/ods', to: 'spreadsheets#ods'
  get 'spreadsheets/xlsx', to: 'spreadsheets#xlsx'
  get 'spreadsheets/alt_xlsx', to: 'spreadsheets#alt_xlsx'
end
