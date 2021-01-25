#http://d.hatena.ne.jp/keyesberry/touch/20111229/p1
#http://qiita.com/akihyro/items/432f63ad9dc90f415e2d
#
# Excel 列名⇔数字
#
def alpha2num(alphabets)
 [*'A'..alphabets].size
end

def num2alpha(number)
 alpha = 'A'
 (number-1).times {alpha.succ!}
 alpha
end

alphabets = %w(A B Z AA AB AZ BB AAA IV ZZZ XFD)
numbers = alphabets.map { |alpha| alpha2num alpha }
  # => [1, 2, 26, 27, 28, 52, 54, 703, 256, 18278, 16384]
numbers.map { |num| num2alpha num }
  # => ["A", "B", "Z", "AA", "AB", "AZ", "BB", "AAA", "IV", "ZZZ", "XFD"]

# http://blade.nagaokaut.ac.jp/cgi-bin/scat.rb/ruby/ruby-list/50163
def rgb(red,green,blue)
  # VBAの場合は rrggbb の並びではない
  # http://www.relief.jp/itnote/archives/000482.php
  red | (green << 8) | (blue << 16)
end

class Sheet < Excel # to use excel Union method, need excel win32ole object

  attr_reader :name

  def initialize(sheet)
    @sheet = sheet
    @name  = sheet.name

    @sheet.Cells.Font.Name = "Meiryo UI"
    @sheet.Cells.Font.Size = 10
  end

  def acol(value)
    # return column alphabaet name of value
    ecol   = @sheet.UsedRange.Columns.count
    values = @sheet.Range("A1",@sheet.Cells(1,ecol)).Value.flatten
    tgtcol = values.zip(1..,'A'..).assoc(value) # [["Key",1,"A"],...
    if tgtcol
      tgtcol.last
    else
      $mylogger.error "#{value} is not found in #{@sheet.name}!"
      nil
    end
  end

  def ncol(value)
    # return column number of value
    ecol   = @sheet.UsedRange.Columns.count
    values = @sheet.Range("A1",@sheet.Cells(1,ecol)).Value.flatten
    tgtcol = values.zip(1..,'A'..).assoc(value) # [["Key",1,"A"],...
    if tgtcol
      tgtcol[1]
    else
      $mylogger.error "#{value} is not found in #{@sheet.name}!"
      nil
    end
  end

end
