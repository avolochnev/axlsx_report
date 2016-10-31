require 'axlsx'
require_relative 'column_name_conv'
require_relative 'group'
require_relative 'column'

module AxlsxReport
  class Base
    include ColumnNameConv

    # Creates new report.
    # By default creates own Axlsx::Package.
    # @param [Axlsx::Package] package - provide existing package to add report sheet to combined document.
    def initialize(package = nil)
      @package = package || Axlsx::Package.new
    end

    # Define report column.
    #
    # Usage:
    #   Human = Struct.new(:first_name, :last_name, :birthday)
    #   class Report < AxlsxReport::Base
    #     # calculate column value with block:
    #     column 'First Name', width: 10 do |human|
    #       human.first_name
    #     end
    #     # or lambda in context of given object
    #     column 'Last Name', -> { last_name }, width: 15
    #     # or chagne context using with:
    #     column 'Age', -> { Date.today.year - year }, with: :birthday
    #   end
    #
    # @see Column#initialize for options description
    #
    # @param [String] name - column header
    # @param [Proc] callable - use proc or lambda instead of block to get cell value. Optional.
    # @param [Hash] options - last param,
    # @param block
    # args: callable = nil, options = {}
    def self.column(col_name, *args, &block)
      @columns ||= []
      options =
        if args.last.is_a? Hash
          args.pop
        else
          {}
        end
      options[:group] = @current_group if @current_group
      callable = args.first || block
      @columns << Column.new(col_name, callable, options)
    end

    # Define group of columns
    #
    # Usage:
    #   Human = Struct.new(:first_name, :last_name, :birthday)
    #   class Report < AxlsxReport::Base
    #     group 'Name' do
    #       column 'First', &:first_name
    #       column 'Last',  &:last_name
    #     end
    #     column 'Birthday', &:birthday
    #   end
    #
    # @param [String] name group name
    def self.group(name, &block)
      @columns ||= []
      unless @groups
        @groups = []
        define_method :groups do
          self.class.instance_eval { @groups }
        end
      end
      raise 'Nested groups are not implemented yet' if @current_group
      @current_group = Group.new(name, @columns.length)
      @groups << @current_group
      instance_exec(&block)
      @current_group.end_index = @columns.length - 1
      @current_group = nil
    end

    # Define group row height.
    # Used if any groups are defiend. Autoheight is used by default
    # @param [Integer] num group row height
    def self.group_height(num)
      define_method(:group_height) { num }
    end

    # Define head row height.
    # Autoheight is used by default
    # @param [Integer] num head row height
    def self.head_height(num)
      define_method(:head_height) { num }
    end

    # Add 1,2,3...N row below head row
    def self.numerate_columns
      define_method(:numerate_columns) { true }
    end

    # Define sheet name.
    #
    # @param [String] name (nil) static sheet name.
    # @param block dinamic sheet name (if given)
    def self.sheet_name(name = nil, &block)
      if block_given?
        define_method(:sheet_name, &block)
      else
        define_method(:sheet_name) { name }
      end
    end

    # Add row to the report.
    # @param [Any] obj object providing the data for columns
    def <<(obj)
      row = columns.map { |column| column.value(self, obj) }
      add_totals row
      sheet.add_row row, row_options
    end

    # Returns sheet name.
    #
    # May be overriden to provide custom sheet name
    # @see Base.sheet_name
    def sheet_name
      'Sheet1'
    end

    # def run(file = nil, enum = nil)
    #   file = file_name unless file
    #   enum = enumerator if enum.nil? && !block_given?
    #   if enum
    #     enum.each do |obj|
    #       self << obj
    #     end
    #   elsif block_given?
    #     yield self
    #   else
    #     raise 'No block or enum given. Use #new -> << -> #save for generating report later'
    #   end
    #   save(file)
    # end

    # def self.[](*args)
    #   new.run(*args)
    # end

    # Save xlsx file with provided data
    # @param [String] file File name the data to be serialized
    def save(file)
      done
      begin
        @package.serialize(file)
      rescue Errno::EACCES => e
        puts "#{file} is protected!"
        file = file + ".tmp"
        @package.serialize(file)
      end
    end

    # Finalize document formatting after data collecting and before saving.
    # @note Called during #save.
    # @note Don't call twice. Don't call before all data provided.
    def done
      apply_width(sheet)
      merge_same(sheet)
      sheet.add_row(@totals) if @totals
    end

    # Axlsx sheet.
    # @return [Axlsx::Worksheet]
    def sheet
      @sheet ||= init_sheet(sheet_name)
    end

    private

    def columns
      self.class.instance_eval { @columns }
    end

    def init_sheet(name = 'Sheet1')
      sheet = @package.workbook.add_worksheet(:name => sheetify_name(name))
      if self.respond_to? :title
        if title.is_a? Array
          title.each do |t|
            sheet.add_row [t]
          end
        else
          sheet.add_row [title]
        end
      end
      if defined? groups
        add_group_head sheet
      else
        add_head sheet, *columns.map{ |col| eval_name(col.name) }
        sheet.rows.last.height = head_height if defined? head_height
      end
      if defined? numerate_columns
        add_frozen_head sheet, *(1..(columns.length)).to_a
      else
        froze(sheet, sheet.rows.count)
        auto_filter(sheet)
      end
      @header_rows = sheet.rows.count
      sheet
    end

    def auto_filter(sheet)
      sheet.auto_filter = Axlsx::cell_range(sheet.rows.last.cells, false)
    end

    def apply_width(sheet)
      sheet.column_widths *columns.map(&:width).map { |width| width || default_column_width }
    end

    def merge_same(sheet)
      columns.map(&:options).each_with_index do |ops, col_num|
        next unless ops[:merge_same]
        merged_style ||= sheet.styles.add_style(:alignment => { :vertical => :center })
        col_char = column_num_to_name(col_num + 1)
        last, start_row = nil, 0
        sheet.rows.dup.each_with_index do |row, row_num|
          next if row_num < @header_rows
          col_value = row.cells[col_num].value
          if last && last != col_value && (row_num - start_row) > 1
            range = "#{col_char}#{start_row + 1}:#{col_char}#{row_num}"
            sheet.merge_cells(range)
            sheet[range].each { |c| c.style = merged_style }
          end
          if last != col_value
            last = col_value
            start_row = row_num
          end
        end
        if last && (sheet.rows.length - start_row) > 1
          range = "#{col_char}#{start_row + 1}:#{col_char}#{sheet.rows.length}"
          sheet.merge_cells(range)
          sheet[range].each { |c| c.style = merged_style }
        end
      end
    end

    def default_column_width
      16
    end

    def add_group_head(sheet)
      head_row = *columns.map{ |col| eval_name(col.name) }
      groups.each { |g| head_row[g.start_index] = g.name }
      add_head sheet, *head_row
      group_row = sheet.rows.length
      groups.map { |g| "#{column_num_to_name(g.start_index + 1)}#{group_row}:#{column_num_to_name(g.end_index + 1)}#{group_row}" }
        .each { |rng| sheet.merge_cells(rng) }
      sheet.rows.last.height = group_height if defined? group_height
      add_head sheet, *columns.map{ |col| eval_name(col.name) }
      sheet.rows.last.height = head_height if defined? head_height
      names_row = sheet.rows.length
      units = *columns.map(&:units)
      has_units = !units.all?(&:nil?)
      if has_units
        add_head sheet, *units
      end
      units_row = sheet.rows.length
      columns.map(&:options).each_with_index do |ops, index|
        row_name = column_num_to_name(index + 1)
        case
        when ops[:units] && ops[:group]
          nil
        when ops[:group]
          sheet.merge_cells("#{row_name}#{names_row}:#{row_name}#{units_row}") if has_units
        when ops[:units]
          sheet.merge_cells("#{row_name}#{group_row}:#{row_name}#{names_row}")
        else
          sheet.merge_cells("#{row_name}#{group_row}:#{row_name}#{units_row}")
        end
      end
    end

    def add_totals(row)
      return unless total_actions
      @totals ||= [nil] * columns.length
      total_actions.each_with_index do |action, index|
        case action
          when String
            @totals[index] = action
          when :sum
            if row[index]
              @totals[index] ||= 0.0
              @totals[index] += row[index]
            end
          when :count
            if row[index]
              @totals[index] ||= 0
              @totals[index] += 1
            end
          when Proc
            @totals[index] = action[row]
        end
      end
    end

    def total_actions
      @total_actions ||=
        columns.map(&:options).map do |ops|
          ops[:total]
        end
    end

    def eval_name(name)
      name.is_a?(Proc) ? instance_exec(&name) : name
    end

    def sheetify_name(name)
      name.gsub('/', '_')[0..30]
    end

    def add_head(sheet, *header)
      row = nil
      sheet.workbook.styles do |s|
        head_style = s.add_style :border => { :style => :thin, :color => "00" },
                                 :alignment => { :horizontal => :center,
                                                 :vertical => :center ,
                                                 :wrap_text => true}
        row = sheet.add_row header, style: head_style
      end
      row
    end

    def add_frozen_head(sheet, *header)
      row = add_head(sheet, *header)
      froze(sheet, sheet.rows.count)
      sheet.auto_filter = Axlsx::cell_range(row.cells, false)
    end

    def separate_sheet(sheet, sheet_name = 'Sheet2', &block)
      sheet.workbook.add_worksheet(name: sheet_name) { |sheet| yield sheet }
    end

    def froze(sheet, num_rows = 2)
      sheet.sheet_view do |vs|
        vs.pane do |pane|
          pane.state = :frozen
          pane.y_split = num_rows
        end
      end
    end

    def row_options
      options = {}
      types = columns.map { |c| c.options[:type] }
      options[:types] = types unless types.all?(&:nil?)
      options
    end
  end
end