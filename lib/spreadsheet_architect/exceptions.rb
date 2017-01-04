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

  end
end
