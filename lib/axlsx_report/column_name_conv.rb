module AxlsxReport
  module ColumnNameConv
    extend self

    # 'Z' => 25
    # 'AA' => 26
    def column_name_to_num(name)
      num = 0
      name.upcase.each_char do |c|
        num *= 26 if num > 0
        add = c.ord - "A".ord + 1
        raise "Invalid symbol in Excel column name: '#{c}'" if add < 1 || add > 26
        num += add
      end
      num - 1
    end

    # 1 => 'А'
    # 26 => 'Z'
    # 27 => 'AA'
    # 28 => 'АB'
    def column_num_to_name(num)
      name = ""
      while num > 0
        nm = (num - 1) % 26
        num = (num - 1) / 26
        name << ("A".ord + nm).chr
      end
      name.reverse
    end
  end
end