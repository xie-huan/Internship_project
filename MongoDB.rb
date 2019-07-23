# encoding:gbk
require 'mongo'
require 'singleton'

class MongoDB
  include Singleton

  #防止修改表格式
  attr_reader :test_report_item

  def init(test_report_item,arg)
    # 获取表格格式、连接用户名和密码
    @test_report_item=test_report_item
    user = arg[0]
    password = arg[1]

    # 连接MongoDB数据库
    @client = Mongo::Client.new([ '127.0.0.1:27017'],
                                :database=>"test_report")


    @client.database.users.create(
        user,
        password: password,
        roles: [ Mongo::Auth::Roles::READ_WRITE ])


    @client = @client.with(:database=>"test_report",
                           user:user,
                           password:password)
  end

  def record(collection_name, *arg)
    if check_available(collection_name,arg) == 1
      # 创建插入的文档
      document = {}
      i = 0
      # 获取collection对应的键
      attrs = @test_report_item[collection_name.to_sym]
      attrs.each do |attr|
        document[attr.to_sym]=arg[i]
        i = i + 1
      end

      # 将文档插入MongoDB对应的collection
      @client[collection_name.to_sym].insert_one(document)
    else
      p "插入文档失败，请重新检测参数名和数量后再进行记录。"+"失败记录："+collection_name+"（表名）；"+arg.to_s+"（数据）"
    end
  end

  def check_available(collection_name, arg)
    # 判断用户输入的集合名是否合法，并设定标志
    collection_name_flag = 0
    @test_report_item.keys.each do |key|
      if key.to_s == collection_name
        collection_name_flag = 1
      end
    end

    # 获取collection对应的键
    attrs = @test_report_item[collection_name.to_sym]

    # 判断用户输入的记录是否合法，并设定标志
    arg_flag = 0
    if collection_name_flag and arg != []
      if arg.length == attrs.length
        arg_flag = 1
      end
    end

    if collection_name_flag == 1 and arg_flag == 1
      return 1
    end
    return 0
  end

  def close
    @client.close
  end

  # 不允许外界调用
  private :check_available
end