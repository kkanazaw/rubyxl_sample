# ActiveRecordのSQLログを出力する
ActiveRecord::Base.logger = Logger.new(STDOUT)
 
# RailsLoggerのログを出力する
Rails.logger = Logger.new(STDOUT)
 
# FactoryGirl の記述を省略する
include FactoryGirl::Syntax::Methods
 
# rspec の should_recieve などを利用する
require 'rspec/mocks/standalone'
