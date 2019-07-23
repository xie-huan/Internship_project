# encoding:gbk
require 'win32ole'
require 'singleton'
require_relative '../ds_comm_fun'
class Excel
  include Singleton
  #测试报告路径
  TEST_REPORT_PATH = File.join(File.dirname(__FILE__),'/')

  #报告文件对象
  attr_accessor :excel
  attr_accessor :wb

  #防止修改表格式
  attr_reader :test_report_item

  def init(test_report_item)
    # 设置表头
    @test_report_item=test_report_item
    # 创建路径
    run_time = Time.now.strftime("%Y_%m_%d_%H_%M_%S")
    Dir.mkdir(TEST_REPORT_PATH) unless File.exist?(TEST_REPORT_PATH)
    create(File.join(TEST_REPORT_PATH,"upgrade_test_report_#{run_time}.xlsx").encode('gbk').gsub!(/[\/]/,'\\'))
  end

  #创建测试报告
  def create(path)
    @excel = WIN32OLE.new("excel.application")
    #创建excel文件
    @wb = excel.Workbooks.Add

    @test_report_item.keys.each {|sheet_name|
      #创建sheet
      @wb.worksheets.Add.Name = sheet_name.to_s

      #创建表头
      col = 1
      @test_report_item[sheet_name].each {|col_name|
        ws = @wb.Worksheets(sheet_name.to_s)
        ws.select

        col_letter = DsCommFun::number2letter(col)
        ws.range("#{col_letter}1").value = col_name
        #文字加粗
        ws.range("#{col_letter}1").font.bold = true
        #设置背景颜色
        ws.range("#{col_letter}1").Interior.ColorIndex = 15

        col = col + 1
      }
    }

    #删除Sheet1
    ws = @wb.Worksheets("Sheet1")
    ws.Delete

    #保存表格
    @wb.SaveAs(path)
  end

  #记录测试结果
  def record(sheet_name, *arg)
    ws = @wb.Worksheets(sheet_name)
    ws.select
    row = ws.usedrange.rows.count
    col = ws.usedrange.columns.count

    #容错
    if arg.length == col
      ws.Range("A#{row + 1}:#{DsCommFun::number2letter(col)}#{row + 1}").Value = arg

      #为errlog设置超链接
      if sheet_name == "Detail" and arg[col - 1] =~ /\\/
        ws.hyperlinks.add(ws.Range("#{DsCommFun::number2letter(col)}#{row + 1}"),arg[col - 1])
      end
    end

    #保存
    @wb.Save
  end

  def close()
    #释放wb资源
    if @wb
      @wb.close(0)
      @wb.ole_free
    end
    #释放excel资源
    if excel
      @excel.Quit
      @excel.ole_free
    end
  end

  private :create
end