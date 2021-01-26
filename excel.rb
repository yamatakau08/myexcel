# http://rubyonwindows.blogspot.jp/2008/01/rubygarden-archives-scripting-excel.html
# https://docs.ruby-lang.org/ja/latest/method/WIN32OLE/s/const_load.html
# explains Excel const e.g xlDown,xlToLeft...
module ExcelConst end

USE_MSOTHEMECOLOR = false
# http://blade.nagaokaut.ac.jp/cgi-bin/scat.rb/ruby/ruby-list/50164
# https://docs.microsoft.com/ja-jp/office/vba/api/overview/library-reference/enumerations-office
# arg1 is タイプライブラリ名 Excel VBA を開き
# ツール -> 参照設定 で、表示されるダイログの 参照可能な ライブラリファイル(A) で知る事が可能
module Office end

class Excel

  @@excel = WIN32OLE.new('Excel.Application')
  WIN32OLE.const_load(@@excel, ExcelConst)
  WIN32OLE.const_load('Microsoft Office 16.0 Object Library', Office)

  def initialize
  end

  def open_book(book = nil)
    if book
      # later implement
      warn "#{self.class.name}##{__method__} Need to implement"
      nil
    else
      # http://officetanaka.net/excel/vba/file/file02.htm
      if file = @@excel.Application.GetOpenFilename("Microsoft Excel,*.xls?")
        @book = @@excel.workbooks.open(file)
      else
        warn "#{self.class.name}##{__method__} file is not specified!"
        nil
      end
    end
  end

  def visible
    @@excel.visible = true

    #
    @@excel.Application.WindowState = ExcelConst::XlNormal

    # Excel 表示位置
    @@excel.Application.Top    = 100
    @@excel.Application.Left   = 10
    @@excel.Application.Width  = 640
    @@excel.Application.Height = 640
  end

  def quit
    @@excel.quit
  end

  def self.copy_chartarea
    @@excel.Application.ActiveChart.ChartArea.Copy
  end

  def self.activewindow
    @@excel.ActiveWindow
  end

  def self.union(range1,range2,*range3_30)
    if range3_30.size >= 1
      warn "#{self.class.name}##{__method__} need to implement for #{range3_30}"
    else
      @@excel.Union(range1,range2)
    end
  end
end
