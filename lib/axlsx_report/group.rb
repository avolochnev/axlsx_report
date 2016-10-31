module AxlsxReport
  class Group
    attr_reader :name, :start_index
    attr_accessor :end_index

    def initialize(name, start_index)
      @name = name
      @start_index = start_index
    end
  end
end