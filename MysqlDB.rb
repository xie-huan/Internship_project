
require 'singleton'
require "mysql2"

class MysqlDB
  include Singleton

  #防止修改表格式
  attr_reader :test_report_item

  def init(test_report_item,arg)
    # 获取表格格式、连接用户名和密码
    @test_report_item=test_report_item
    user = arg[0]
    password = arg[1]

    # 新建连接对象
    @client = Mysql2::Client.new(
      :host=>"localhost",
      :username=>user,
      :password=>password,
      #:connect_timeout=>10,
      :init_command=>"create database if not exists report;"
    )
    @client.query("use report;")

    # 新建数据表，如果存在数据表，则创建，否则不创建
    @test_report_item.keys.each { |table_name|
      # 最终 sql 语句的形式：create table if not exists CheckItem(Checker varchar(255),CheckItem varchar(255));
      # 由于表的属性不知道具体，所以选择varchar(255)，可更改，避免浪费空间
      sql = "create table if not exists "+table_name.to_s+"("
      len = @test_report_item[table_name].length
      i = 0
      @test_report_item[table_name].each {|attr|
        i += 1
        if i==len
          sql += attr+" varchar(255)"
        else
          sql += attr+" varchar(255),"
        end
      }
      sql +=");"
      @client.query(sql)
    }
  end

  #记录测试结果
  def record(sheet_name, *arg)
    #创建插入sql语句：insert into checkitem values("1","2");
    insert_sql = "insert into "+sheet_name+' values("'
    len = arg.length
    i=0
    arg.each do |v|
      i+=1
      if i == len
        insert_sql+=v+'"'
      else
        insert_sql+=v+'","'
      end
    end
    insert_sql+=");"

    # 对文件路径进行处理
    if sheet_name == "Detail" and arg[len - 1] =~ /\\/
      insert_sql.gsub!(/\\/,'\\\\\\')
    end

    # 执行插入语句
    @client.query(insert_sql)
  end

  def close
    @client.close
  end
end