Rails.application.routes.draw do
  get '/contact', to: 'contact#new'
  post '/contact', to: 'contact#create'
 

  get 'generate_excel', to: 'contact#generate_excel'   
  

 
  get 'contact/download_excel', to: 'contact#download_excel', as: 'download_excel'


   root "contact#new"

end
