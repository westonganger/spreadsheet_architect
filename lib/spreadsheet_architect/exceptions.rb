module SpreadsheetArchitect
  module Exceptions

    class NoDataError < StandardError
      def initialize
        super("Missing :data or :instances option")
      end
    end

    class MultipleDataSourcesError < StandardError
      def initialize
        super("Both :data and :instances options cannot be combined, please choose one.")
      end
    end

    class IncorrectTypeError < StandardError
      def initialize(option=nil)
        super("Incorrect data type for :#{option} option")
      end
    end
    
    class SpreadsheetColumnsNotDefinedError < StandardError
      def initialize(klass=nil)
        super("The instance method `spreadsheet_columns` is not defined on #{klass.name}")
      end
    end

    class InvalidColumnError < StandardError
      def initialize(col)
        super("Invalid Column `#{col}` given for column_types options")
      end
    end

    class InvalidRangeStylesOptionError < StandardError
      def initialize(type, opt)
        super("Invalid or missing :#{type} option for `#{opt}`. :#{type} can be an integer, range, or :all")
      end
    end

    class BadRangeError < StandardError
      def initialize(type, range)
        case type
        when :columns, :rows
          super("Bad range `#{range}` passed. Some of the #{type} specified were greater than the total number of #{type}")
        when :format
          super("Bad range `#{range}` passed. Format must be as follows: A1:D4")
        when :type
          super("Incorrect range type `#{range}`. Valid types are String and Hash")
        else
          super("Bad range `#{range}` passed.")
        end
      end
    end
  end
end
