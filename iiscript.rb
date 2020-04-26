require 'sinatra'
require 'docx'
require 'date'
require 'mail'



get '/' do

if Time.now.wday != 0

   date = Time.new.day.to_s + '.0' + Time.new.month.to_s + '.' + Time.new.year.to_s
   doc = Docx::Document.open('template.docx')

   doc.paragraphs.each do |p|
    p.each_text_run do |tr|
      tr.substitute('date',"#{date}")
    end
   end

    doc.save("#{Time.new.day.to_s}.docx")

  #Mail
    Mail.defaults	 do
      delivery_method :smtp, {
        address: 'smtp.mail.ru',
        port: 587,
        domain: 'mail.ru',
        user_name: 'artik3314@mail.ru',
        password: 'Aa003314+',
        authentication: :plain,
        enable_starttls_auto: true
      }
    end

    Mail.deliver do
      from      "artik3314@mail.ru"
      to        "irinaluk69@mail.ru"
      subject   "Служебная записка"
      body      "Мониторинг П 172"
      add_file  "#{Time.new.day.to_s}.docx"
    end

  erb :index
 else
  erb :wday
end

end



=begin

  #Gmail
  Mail.defaults do
    delivery_method :smtp, {
      address: 'smtp.gmail.com',
      port: 587,
      domain: 'gmail.com',
      user_name: 'artik3314@gmail.com',
      password: '00003314',
      authentication: :plain,
      enable_starttls_auto: true
    }
  end

  	Mail.deliver do
	  from      "artik3314@gmail.com"
	  to        "artik3314@gmail.com"
	  subject   ""
	  body      ""
	  add_file  "#{Time.new.day.to_s}.docx"
	end	
  
#Yandex
  Mail.defaults	 do
    delivery_method :smtp, {
      address: 'smtp.yandex.ru',
      port: 587,
      domain: '127.0.0.1:3000',
      user_name: 'artik3314@yandex.ru',
      password: '003314',
      authentication: :plain,
      enable_starttls_auto: true
    }
  end

	Mail.deliver do
	  from      "artik3314@yandex.ru"
	  to        "artik3314@yandex.ru"
	  subject   ""
	  body      ""
	  add_file  "#{Time.new.day.to_s}.docx"
	end

=end


