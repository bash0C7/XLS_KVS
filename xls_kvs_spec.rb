# -*- encoding: sjis -*-

require 'xls_kvs'
require 'win32ole'

describe XLS_KVS, "書き変えて閉じたとき" do
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

  it "件数は書き変えた値のまま" do
    @kvs.size.should == 3
  end

  it "書き変えたまままなのでからっぽではない" do
    @kvs.empty?.should be_false
  end

  it "書き変えたvalueと同じものを取得できる" do
    @kvs[0].should == "a"
    @kvs[1].should == "b"
    @kvs[2].should == "c"
  end

  it "書き変えたkeyが存在する" do
    @kvs.key?(0).should be_true
    @kvs.key?(1).should be_true
    @kvs.key?(2).should be_true
  end
end

describe XLS_KVS, "からっぽの場合" do
  before(:all) do
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    xls_file = @fso.GetAbsolutePathName('xls.xls')
    @kvs = XLS_KVS.load(xls_file, 1)
    @kvs.clear
  end

  after(:all) do
    @kvs.close
  end

   it "何もしなけりゃからっぽ" do
    @kvs.empty?.should be_true
  end

  it "何もしなければ0件" do
    @kvs.size.should == 0
  end

  it "終了させたら空のまま" do
    @kvs.clear
  end
end

describe XLS_KVS, "存在しないキーを3件追加したとき" do
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
 
  it "件数は追加した分増える" do
    @kvs.size.should == 3
  end

  it "からっぽではなくなる" do
    @kvs.empty?.should be_false
  end

  it "入れたvalueと同じものを取得できる" do
    @kvs[0].should == "a"
    @kvs[1].should == "b"
    @kvs[2].should == "c"
  end

  it "いれたkeyが存在する" do
    @kvs.key?(0).should be_true
    @kvs.key?(1).should be_true
    @kvs.key?(2).should be_true
  end
end

describe XLS_KVS, "存在するキーの値を上書きしたとき" do
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

   it "件数は変わらない" do
    @kvs.size.should == 1
  end

  it "からっぽではない" do
    @kvs.empty?.should be_false
  end

  it "上書きした値を取得できる" do
    @kvs[0].should == "replace"
  end
end

describe XLS_KVS, "存在するキー3件のうち1件を消した時" do
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

   it "件数は消した分減る" do
    @kvs.size.should == 2 
  end

  it "からっぽではない" do
    @kvs.empty?.should be_false
  end

  it "消したキーは取得できない" do
    @kvs[0].should be_nil
    @kvs[1].should == "b"
    @kvs[2].should == "c"
  end

  it "消したキーは存在しない" do
    @kvs.key?(0).should be_false
    @kvs.key?(1).should be_true
    @kvs.key?(2).should be_true
  end
end

#  it "Sheet not found => IOError" do
#    lambda {XLS.foreach(@xls_file, 0)}.should raise_error(IOError)
#    lambda {XLS.foreach(@xls_file, 'nothing')}.should raise_error(IOError)
#  end
