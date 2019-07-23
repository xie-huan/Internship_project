
require 'singleton'
require "erb"

class HTML
  include Singleton

  # 表名数组 和 属性数组
  attr_accessor :table_name
  # 数据存储
  attr_accessor :data
  #防止修改表格式
  attr_reader :test_report_item

  def init(test_report_item)
    @test_report_item = test_report_item

    # 创建html文件存放路径
    run_time = Time.now.strftime("%Y_%m_%d_%H_%M_%S")
    @path="./html/test_report_#{run_time}"
    Dir.mkdir(@path) unless File.exist?(@path)

    # 初始化对象变量，确定对象变量类型
    # table_name存放所有表名
    # data存放所有表名及其对应的数据
    @table_name=[]
    @data = {}

    # 生成每个HTML页面，格式为表格
    @test_report_item.keys.each { |table_name|
      @table_name.push(table_name)

      # attrs在erb模板中会使用（不能删）
      attrs = @test_report_item[table_name]
      html_name = "#{@path}/#{table_name}.html"
      html_in = ""
      File.open("./html/erb_template/format.html.erb",'r') do |file|
        html_in = file.read
      end
      renderer = ERB.new(html_in)
      html_out = renderer.result(binding) #动态绑定
      #写入html文档
      File.open(html_name, 'w') do |file|
        file << html_out
      end
    }

    # 生成导航页面，可导航到每个数据表
    html_in = ""
    File.open("./html/erb_template/navigation.html.erb", 'r') do |file|
      html_in = file.read
    end
    renderer = ERB.new(html_in)
    html_out = renderer.result(binding) #动态绑定
    # 写入html文档
    File.open("#{@path}/Navigation.html", 'w') do |file|
      file << html_out
    end
  end

  # 写入记录
  def record(table_name, *arg)
    if check_available(table_name,arg) == 1
      # 将记录写到对象变量，然后写到HTML页面
      if data[table_name] == nil
        data[table_name]=[]
        @data[table_name].push(arg)
      else
        @data[table_name].push(arg)
      end

      #重新生成对应记录表格的html
      attrs = @test_report_item[table_name.to_sym]
      html_in = ""
      File.open("./html/erb_template/data.html.erb", 'r') do |file|
        html_in = file.read
      end
      renderer = ERB.new(html_in)
      html_out = renderer.result(binding) #动态绑定
      # 写入html文档
      File.open("#{@path}/#{table_name}.html", 'w') do |file|
        file << html_out
      end
    else
      p "记录失败"
    end

  end

  def check_available(table_name, arg)
    # 判断用户输入的集合名是否合法，并设定标志
    table_name_flag = 0
    @test_report_item.keys.each do |key|
      if key.to_s == table_name
        table_name_flag = 1
      end
    end

    # 获取collection对应的键
    attrs = @test_report_item[table_name.to_sym]

    # 判断用户输入的记录是否合法，并设定标志
    arg_flag = 0
    if table_name_flag and arg != []
      if arg.length == attrs.length
        arg_flag = 1
      end
    end

    if table_name_flag == 1 and arg_flag == 1
      return 1
    end
    return 0
  end

  # 关闭文件
  def close
    #@html_file.close
    system("start #{@path}/navigation.html")
  end

  private :check_available
end
