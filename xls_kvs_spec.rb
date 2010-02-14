# -*- encoding: sjis -*-

require 'xls_kvs'
require 'win32ole'

describe XLS_KVS, "�����ς��ĕ����Ƃ�" do
  before(:all) do
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    xls_file = @fso.GetAbsolutePathName('xls.xls')
    @kvs = XLS_KVS.load(xls_file, 1)
    @kvs.clear
    @kvs.store(0, "a")
    @kvs.store(1, "b")
    @kvs.store(2, "c")
    @kvs.close
    @kvs = XLS_KVS.load(xls_file, 1)
  end

  after(:all) do
    @kvs.close
  end

  it "�����͏����ς����l�̂܂�" do
    @kvs.size.should == 3
  end

  it "�����ς����܂܂܂Ȃ̂ł�����ۂł͂Ȃ�" do
    @kvs.empty?.should be_false
  end

  it "�����ς���value�Ɠ������̂��擾�ł���" do
    @kvs[0].should == "a"
    @kvs[1].should == "b"
    @kvs[2].should == "c"
  end

  it "�����ς���key�����݂���" do
    @kvs.key?(0).should be_true
    @kvs.key?(1).should be_true
    @kvs.key?(2).should be_true
  end
end

describe XLS_KVS, "������ۂ̏ꍇ" do
  before(:all) do
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    xls_file = @fso.GetAbsolutePathName('xls.xls')
    @kvs = XLS_KVS.load(xls_file, 1)
    @kvs.clear
  end

  after(:all) do
    @kvs.close
  end

   it "�������Ȃ���Ⴉ�����" do
    @kvs.empty?.should be_true
  end

  it "�������Ȃ����0��" do
    @kvs.size.should == 0
  end

  it "�I�����������̂܂�" do
    @kvs.clear
  end
end

describe XLS_KVS, "���݂��Ȃ��L�[��3���ǉ������Ƃ�" do
  before(:all) do
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    xls_file = @fso.GetAbsolutePathName('xls.xls')
    @kvs = XLS_KVS.load(xls_file, 1)
    @kvs.clear
    @kvs.store(0, "a")
    @kvs.store(1, "b")
    @kvs.store(2, "c")

  end
 
  after(:all) do
    @kvs.close
  end
 
  it "�����͒ǉ�������������" do
    @kvs.size.should == 3
  end

  it "������ۂł͂Ȃ��Ȃ�" do
    @kvs.empty?.should be_false
  end

  it "���ꂽvalue�Ɠ������̂��擾�ł���" do
    @kvs[0].should == "a"
    @kvs[1].should == "b"
    @kvs[2].should == "c"
  end

  it "���ꂽkey�����݂���" do
    @kvs.key?(0).should be_true
    @kvs.key?(1).should be_true
    @kvs.key?(2).should be_true
  end
end

describe XLS_KVS, "���݂���L�[�̒l���㏑�������Ƃ�" do
  before(:all) do
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    xls_file = @fso.GetAbsolutePathName('xls.xls')
    @kvs = XLS_KVS.load(xls_file, 1)
    @kvs.clear
    @kvs.store(0, "a")
    @kvs.store(0, "replace")
  end

  after(:all) do
    @kvs.close
  end

   it "�����͕ς��Ȃ�" do
    @kvs.size.should == 1
  end

  it "������ۂł͂Ȃ�" do
    @kvs.empty?.should be_false
  end

  it "�㏑�������l���擾�ł���" do
    @kvs[0].should == "replace"
  end
end

describe XLS_KVS, "���݂���L�[3���̂���1������������" do
  before(:all) do
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    xls_file = @fso.GetAbsolutePathName('xls.xls')
    @kvs = XLS_KVS.load(xls_file, 1)
    @kvs.clear
    @kvs.store(0, "a")
    @kvs.store(1, "b")
    @kvs.store(2, "c")
    @kvs.delete(0)
  end

  after(:all) do
    @kvs.close
  end

   it "�����͏�����������" do
    @kvs.size.should == 2 
  end

  it "������ۂł͂Ȃ�" do
    @kvs.empty?.should be_false
  end

  it "�������L�[�͎擾�ł��Ȃ�" do
    @kvs[0].should be_nil
    @kvs[1].should == "b"
    @kvs[2].should == "c"
  end

  it "�������L�[�͑��݂��Ȃ�" do
    @kvs.key?(0).should be_false
    @kvs.key?(1).should be_true
    @kvs.key?(2).should be_true
  end
end

#  it "Sheet not found => IOError" do
#    lambda {XLS.foreach(@xls_file, 0)}.should raise_error(IOError)
#    lambda {XLS.foreach(@xls_file, 'nothing')}.should raise_error(IOError)
#  end
