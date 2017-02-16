module SpreadsheetArchitect
  module Exceptions

    class NoDataError < StandardError
      def initialize
        super("Missing :data option")
      end
    end

    class NoInstancesError < StandardError
      def initialize
        super("Missing :instances option")
      end
    end

    class IncorrectTypeError < StandardError
      def initialize(option=nil)
        super("Incorrect data type for :#{option} option")
      end
    end
    
    class SpreadsheetColumnsNotDefinedError < StandardError
      def initialize(klass=nil)
        super("The spreadsheet_columns option is not defined on #{klass.name}")
      end
    end

    class InvalidColumnError < StandardError
      def initialize(range_hash)
        super("Invalid Column `#{range_hash}` given for column_types options")
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
        end
      end
    end
  end
end
