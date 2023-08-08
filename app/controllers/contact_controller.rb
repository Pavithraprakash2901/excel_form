

class ContactController < ApplicationController

  def new
    @contact = Contact.new
  end

  def create
    @contact = Contact.new(contact_params)
    if @contact.save
      respond_to do |format|
        format.html { redirect_to root_path, notice: 'Contact was successfully created.' }
        format.json { render json: { message: 'Contact was successfully created.' } }
        format.js
      end
    else
      respond_to do |format|
        format.html { render :new }
        format.json { render json: { errors: @contact.errors.full_messages }, status: :unprocessable_entity }
        format.js  
      end
    end
  end







  def download_excel
    contacts = Contact.all
  
    workbook = FastExcel.open
    worksheet = workbook.add_worksheet

    format = workbook.add_format(
      italic: true , bold: true ,  align: {h: :align_center, v: :align_vertical_center},bottom_color: :alice_blue)
  
    headers = ['Name', 'Email', 'Date']
    # worksheet.set_row(0, 20, bold_font) # Apply bold formatting to the whole first row (header row)
    worksheet.append_row(headers,format)

    
    date_format = workbook.number_format("[$-409]m/d/yy   h:mm AM/PM;@")
    # worksheet.set_column(2, 2, 20, date_format)

  
    # Add contact data to the worksheet
    contacts.each do |contact|
      row_data = [contact.name, contact.email, contact.date] # Adjust 'date' to the actual attribute name in your Contact model
      worksheet.append_row(row_data)
    end
  
    # Set column widths
    worksheet.set_column(0, 0, 25) # Set width of the first column (Name) to 20
    worksheet.set_column(1, 1, 35) # Set width of the second column (Email) to 30
    worksheet.set_column(2, 2, 20, date_format)
    # Set width of the third column (Date) to 15

   
    # Set the response headers for downloading the file
    filename = 'contact_data.xlsx'
    send_data workbook.read_string, filename: filename, type: 'application/xlsx'
  end
  





  
  private

  def contact_params
    params.require(:contact).permit(:name, :email, :date)
  end

end
