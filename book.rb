class Book
  attr_reader :name

  def initialize(book)
    @book = book
    @name = book.name
  end

  def sheet_exist?(sheet_name)
    # if sheet_name exists, return sheet object
    # worksheets does not have map method
    # @sheets = @book.worksheets.map {|sheet| [sheet.name, sheet_object] }.to_h
    @sheets = {}
    @book.worksheets.each do |sheet_obj|
      @sheets.store(sheet_obj.name,sheet_obj)
    end
    @sheets[sheet_name]
  end

  def sheet_add(sheet_name = nil)
    if sheet_exist?(sheet_name)
      warn "#{self.class.name}##{__method__} #{sheet_name} exists!"
    else
      sheet = @book.worksheets.add
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
    @book.close
  end

end
