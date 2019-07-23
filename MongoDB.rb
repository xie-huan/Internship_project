# encoding:gbk
require 'mongo'
require 'singleton'

class MongoDB
  include Singleton

  #��ֹ�޸ı��ʽ
  attr_reader :test_report_item

  def init(test_report_item,arg)
    # ��ȡ����ʽ�������û���������
    @test_report_item=test_report_item
    user = arg[0]
    password = arg[1]

    # ����MongoDB���ݿ�
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
      # ����������ĵ�
      document = {}
      i = 0
      # ��ȡcollection��Ӧ�ļ�
      attrs = @test_report_item[collection_name.to_sym]
      attrs.each do |attr|
        document[attr.to_sym]=arg[i]
        i = i + 1
      end

      # ���ĵ�����MongoDB��Ӧ��collection
      @client[collection_name.to_sym].insert_one(document)
    else
      p "�����ĵ�ʧ�ܣ������¼����������������ٽ��м�¼��"+"ʧ�ܼ�¼��"+collection_name+"����������"+arg.to_s+"�����ݣ�"
    end
  end

  def check_available(collection_name, arg)
    # �ж��û�����ļ������Ƿ�Ϸ������趨��־
    collection_name_flag = 0
    @test_report_item.keys.each do |key|
      if key.to_s == collection_name
        collection_name_flag = 1
      end
    end

    # ��ȡcollection��Ӧ�ļ�
    attrs = @test_report_item[collection_name.to_sym]

    # �ж��û�����ļ�¼�Ƿ�Ϸ������趨��־
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

  # ������������
  private :check_available
end