# XLS_KVS -- key-value store class library using Microsoft Excel.
#        dependent on win32ole & Microsoft Excel
#
#        Programmed by Toshiaki <koshiba+rubyforge@4038nullpointer.com>
#        License : Ruby License

require 'win32ole'
require 'YAML'

class XLS_KVS
  
  #  <description>
  # 
  # ARGS
  # 
  #   <arg> : <description>
  # 
  # RETURNS
  #   <> 
  #   
  # Example
  #

  class << self
    def load(path, sheet, is_readonly = false)
      XLS_KVS::Hash.new(path, sheet, is_readonly)
    end
  end
  
  class Hash
#    include Enumerable
    
    FIRSTITEM = 1
    KEYVALUE_COLS = 'A:B'
    KEY_COLS = 'A:A'

    def Hash.finalize(app, book, is_readonly)
      proc {
        book.save if (book && !(is_readonly))	    
        book.close({'SaveChanges' => !(is_readonly)}) if book
        app.Quit if app
      }
    end    

    def initialize(path, sheet, is_readonly, server_name = nil)
      @app = WIN32OLE.new('Excel.Application', server_name)
      #@app.visible = true
      @book = @app.Workbooks.Open(path,{'ReadOnly' => is_readonly})
      @sheet = @book.sheets(sheet)
      @sheet.activate
      ObjectSpace.define_finalizer(self, Hash.finalize(@app, @book, is_readonly))
      @lock = Mutex.new
    rescue
      close(false)
      raise IOError.new
    end

    def [](key)
      YAML.load(@app.WorksheetFunction.VLookup(YAML.dump(key), 
				     @sheet.Range(KEYVALUE_COLS), 
				     2, 
				     false)
	       )
    rescue
      nil
    end

    def empty?
      size == 0
    end

    def size
      @app.WorksheetFunction.Counta(@sheet.Range(KEY_COLS))
    end

    def clear
      @lock.synchronize {
	    @sheet.UsedRange.clear
      }	    
      self
    end	    

    def delete(key)
      value = nil	    
      @lock.synchronize {
         value = self[key]	
	 target_row = find(key).Row
         @sheet.Range("#{target_row}:#{target_row}").Delete (-4162)
      }	if key?(key)    
      value
    end	    

    def store(key, value)
      case @app.WorksheetFunction.CountIf(@sheet.Range(KEY_COLS),
					YAML.dump(key)).to_i
      when 0
	insert(key, value)      
      when 1
	replace(key, value)      
      else
         raise IOError.new
      end
      value
    end

    def replace(key, value)
      @lock.synchronize {  
        range = find(key)      
        range.offset(0, 1).value = YAML.dump(value) if range
      }
    end	    

    def find(key)
      range = @sheet.range(KEY_COLS).Find(YAML.dump(key),
                             @app.ActiveCell,
                             -4163, #xlValues, 
                             1, #xlWhole, 
                             1, #xlByRows, 
                             1, #xlNext, 
                             true, 
                             false)
      range
    rescue
      nil
    end

    def key?(key)
      find(key) ? true : false
    end 	   

    def insert(key, value)
      @lock.synchronize {
	  max_row = @sheet.UsedRange.Row + @sheet.UsedRange.Rows.count
   	  @sheet.range("A#{max_row}").value = YAML.dump(key)
          @sheet.range("B#{max_row}").value = YAML.dump(value)
      }
    end	    

    def close(is_save = true)
      @book.save if (@book && is_save)	    
      @book.close({'SaveChanges' => is_save}) if @book
      @app.Quit if @app
    end

  end
  
end
