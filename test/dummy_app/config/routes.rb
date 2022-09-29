Dummy::Application.routes.draw do
  get 'spreadsheets/csv', to: 'spreadsheets#csv'
  get 'spreadsheets/ods', to: 'spreadsheets#ods'
  get 'spreadsheets/xlsx', to: 'spreadsheets#xlsx'
  get 'spreadsheets/test_respond_with', to: 'spreadsheets#test_respond_with'
end
