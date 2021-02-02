class Book
  attr_reader :name

  def initialize(book)
    @book = book
    @name = book.name
  end

  def sheet_exist?(sheet_name)
    # if sheet_name exists, return sheet object
    # worksheets does not have map method
    # @sheets = @book.Worksheets.map {|sheet| [sheet.name, sheet_object] }.to_h
    @book.Worksheets.each do |sheet_obj|
      (@sheets ||= {}).store(sheet_obj.name,sheet_obj)
    end

    @sheets[sheet_name]
  end

  def sheet_add(sheet_name = nil)
    if sheet_exist?(sheet_name)
      warn "#{self.class.name}##{__method__} #{sheet_name} exists!"
    else
      sheet = @book.Worksheets.Add
      sheet.name = sheet_name if sheet_name
      sheet
    end
  end

  def sheet_copy(src_sheet_name = nil,copied_sheet_name = nil)
    unless src_sheet = sheet_exist?(src_sheet_name)
      warn "#{self.class.name}##{__method__} #{src_sheet_name} sheet not found!"
    else
      # https://msdn.microsoft.com/JA-JP/library/office/ff837784.aspx
      # 行の最後のsheetの意味は、sheet変数代入時に指定したシート
      # VBAの名前付き変数は、RubyのWin32OLE場合,ハッシュ変数になる
      # http://d.hatena.ne.jp/maluboh/20080704/p1
      src_sheet.copy(Before: src_sheet)

      # sheet copy後、Active Sheetが、copy先のsheetに変わる。
      # シート名、"general_report (2)" " (2)"が追加される
      sheet = @book.activesheet
      sheet.name = copied_sheet_name if copied_sheet_name
      sheet
    end
  end

  def close
    @book.Close
  end

  def fullname
    @book.FullName
  end

  def sheet_activate(sheet_name)
    # https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.activate(method)
    # same as clik sheet tab
    if sheet = sheet_exist?(sheet_name)
      sheet.Activate
      sheet
    else
      warn "#{self.class.name}##{__method__} #{sheet_name} sheet not found!"
    end
  end

  def copy_range_as_picture(sheet_name,range)
    # range e.g. "A1:B28"
    if sheet = sheet_activate(sheet_name)
      # https://www.moug.net/tech/exvba/0050118.html
      #@book.ActiveSheet.Range(range).CopyPicture(Appearance: Excel::XlScreen, Format: Excel::XlPicture)
      sheet.Range(range).CopyPicture(Appearance: Excel::XlScreen, Format: Excel::XlPicture)
    else
      warn "#{self.class.name}##{__method__} #{sheet_name} sheet found!"
    end
  end

  def copy_range(sheet_name,range)
    # range e.g. "A1:B28"
    if sheet = sheet_activate(sheet_name)
      # https://www.moug.net/tech/exvba/0050118.html
      sheet.Range(range).Copy
    else
      warn "#{self.class.name}##{__method__} #{sheet_name} sheet not found!"
    end
  end

  def copy_chart(sheet_name,chart_object_no)
    if sheet = sheet_activate(sheet_name)
      if chart_object_no <= sheet.ChartObjects.Count
        sheet.ChartObjects(chart_object_no).Activate
        Excel.chartareacopy
      else
        warn "#{self.class.name}##{__method__} #{sheet_name} char_object_no: #{chart_object_no} > #{sheet.ChartObjects.Counts}!"
      end
    else
      warn "#{self.class.name}##{__method__} #{sheet_name} sheet not found!"
    end
  end

end
