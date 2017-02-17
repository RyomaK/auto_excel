# coding: UTF-8
  require 'sinatra'
  require 'rubyXL'


get "/" do
	erb :index 
end

post "/excel" do
	@lesson1_1 = params[:lesson1_1]
	@lesson1_2 = params[:lesson1_2]

	xls ="format.xlsx"

	book = RubyXL::Parser.parse(xls) 
	sheet = book["Sheet1"]
	sheet.add_cell(0,1,@lesson1_1)
	sheet.add_cell(0,2,@lesson1_2)


	book.write(xls)

	redirect "/"
end