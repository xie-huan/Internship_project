
require 'singleton'

require_relative 'Excel'
require_relative 'HTML'
require_relative 'MysqlDB'
require_relative 'MongoDB'

class TypeFactory
  include Singleton

  #根据参数类型决定返回哪一个单例
  def create(type,test_report_item,*args)
    if type == "Excel"
      p "Excel"
      Excel.instance.init(test_report_item)
      return Excel.instance
    elsif type == "HTML"
      p "HTML"
      HTML.instance.init(test_report_item)
      return HTML.instance
    elsif type == "MysqlDB"
      p "MysqlDB"
      MysqlDB.instance.init(test_report_item,args)
      return MysqlDB.instance
    elsif type == "MongoDB"
      MongoDB.instance.init(test_report_item,args)
      return MongoDB.instance
    else
      p "Unkonwn type"
    end
  end

end