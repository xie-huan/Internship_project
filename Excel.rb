# encoding:gbk
require 'win32ole'
require 'singleton'
require_relative '../ds_comm_fun'
class Excel
  include Singleton
  #���Ա���·��
  TEST_REPORT_PATH = File.join(File.dirname(__FILE__),'/')

  #�����ļ�����
  attr_accessor :excel
  attr_accessor :wb

  #��ֹ�޸ı��ʽ
  attr_reader :test_report_item

  def init(test_report_item)
    # ���ñ�ͷ
    @test_report_item=test_report_item
    # ����·��
    run_time = Time.now.strftime("%Y_%m_%d_%H_%M_%S")
    Dir.mkdir(TEST_REPORT_PATH) unless File.exist?(TEST_REPORT_PATH)
    create(File.join(TEST_REPORT_PATH,"upgrade_test_report_#{run_time}.xlsx").encode('gbk').gsub!(/[\/]/,'\\'))
  end

  #�������Ա���
  def create(path)
    @excel = WIN32OLE.new("excel.application")
    #����excel�ļ�
    @wb = excel.Workbooks.Add

    @test_report_item.keys.each {|sheet_name|
      #����sheet
      @wb.worksheets.Add.Name = sheet_name.to_s

      #������ͷ
      col = 1
      @test_report_item[sheet_name].each {|col_name|
        ws = @wb.Worksheets(sheet_name.to_s)
        ws.select

        col_letter = DsCommFun::number2letter(col)
        ws.range("#{col_letter}1").value = col_name
        #���ּӴ�
        ws.range("#{col_letter}1").font.bold = true
        #���ñ�����ɫ
        ws.range("#{col_letter}1").Interior.ColorIndex = 15

        col = col + 1
      }
    }

    #ɾ��Sheet1
    ws = @wb.Worksheets("Sheet1")
    ws.Delete

    #������
    @wb.SaveAs(path)
  end

  #��¼���Խ��
  def record(sheet_name, *arg)
    ws = @wb.Worksheets(sheet_name)
    ws.select
    row = ws.usedrange.rows.count
    col = ws.usedrange.columns.count

    #�ݴ�
    if arg.length == col
      ws.Range("A#{row + 1}:#{DsCommFun::number2letter(col)}#{row + 1}").Value = arg

      #Ϊerrlog���ó�����
      if sheet_name == "Detail" and arg[col - 1] =~ /\\/
        ws.hyperlinks.add(ws.Range("#{DsCommFun::number2letter(col)}#{row + 1}"),arg[col - 1])
      end
    end

    #����
    @wb.Save
  end

  def close()
    #�ͷ�wb��Դ
    if @wb
      @wb.close(0)
      @wb.ole_free
    end
    #�ͷ�excel��Դ
    if excel
      @excel.Quit
      @excel.ole_free
    end
  end

  private :create
end