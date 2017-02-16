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
      def initialize
        super("Invalid Column given for column_types options")
      end
    end

    class InvalidRangeStylesOptionError < StandardError
      def initialize(type)
        super("Invalid type for range_styles #{type} option. Allowed formats are Integer, Range, or :all")
      end
    end

    class BadRangeError < StandardError
      def initialize(type)
        case type
        when :columns, :rows
          super("Bad range passed. Some of the #{type} specified were greater than the total number of #{type}")
        when :format
          super('Bad range passed. Format must be as follows: A1:D4')
        when :type
          super('Incorrect range type. Valid types are String and Hash')
        end
      end
    end
  end
end
