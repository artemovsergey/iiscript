require 'sinatra'
#require 'docx'
#require 'date'
#require 'mail'

get '/' do
  erb :index
end

=begin
date = Time.new.day.to_s + '.0' + Time.new.month.to_s + '.' + Time.new.year.to_s

doc = Docx::Document.open('template.docx')


doc.paragraphs.each do |p|
    p.each_text_run do |tr|
      tr.substitute('date',"#{date}")
    end
  end
 
doc.save("#{Time.new.day.to_s}.docx")


=begin
# Retrieve and display paragraphs
doc.paragraphs.each do |p|
  puts p
end

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

=end
