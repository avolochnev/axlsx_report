module AxlsxReport
  # Report's column representation
  class Column
    attr_reader :name, :callable, :options

    # Creates new column
    #
    # @param [String] name Column name
    # @param [Proc] callable Proc to calculate column value form given object
    # @param [Hash] options ({}) Column parameters
    # @option options [Integer] width: (16) column width
    # @option options [Symbol] transform: report object method to be called after value calculation
    # @option options [Symbol] with: method name of given object to be used to base for cell value calculation
    # @option options [String] units: Units text to be added to the header in the column
    # @option options [Symbol] type: axlsx type to be used in this column. One of Axlsx::Cell::CELL_TYPES
    def initialize(name, callable, options = {})
      @name, @callable, @options = name, callable, options
    end

    # Quick access to options
    %i{width units}.each do |option|
      define_method(option) do
        @options[option]
      end
    end

    # Calculates column value for given object
    #
    # @param [AxlsxReport::Base] report Report the value is calculated for
    # @param [Any] obj Source object for column value
    # @return Column value for given object.
    def value(report, obj)
      source = extract_source(obj)
      return nil if source.nil?
      value =
        if callable.arity.zero?
          source.instance_exec &callable
        else
          report.instance_exec source, &callable
        end
      transform = options[:transform]
      value = report.send transform, value if transform
      value
    end

    private

      def extract_source(obj)
        with = options[:with]
        return obj unless with
        return nil unless obj.respond_to?(with)
        obj.send(with)
      end
  end
end