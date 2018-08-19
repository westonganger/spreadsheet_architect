module SpreadsheetArchitect
  module Exceptions

    class OptionTypeError < TypeError
      def initialize(which)
        super("Invalid data type for the #{which} option")
      end
    end

    class ArgumentError < ::ArgumentError
      def initialize(msg)
        super(msg)
      end
    end

    class InvalidRangeError < ArgumentError
      def initialize(type, range)
        case type
        when :columns, :rows
          super("Invalid range `#{range}` passed. Some of the #{type} specified were greater than the total number of #{type}")
        when :format
          super("Invalid range `#{range}` passed. Format must be as follows: A1:D4")
        when :type
          super("Invalid range type `#{range}`. Valid types are String and Hash")
        else
          super("Invalid range `#{range}` passed.")
        end
      end
    end

    class InvalidColumnError < ArgumentError
      def initialize(col)
        super("Invalid Column `#{col}` given for column_types options")
      end
    end

    class InvalidRangeStylesOptionError < ArgumentError
      def initialize(type, opt)
        super("Invalid or missing :#{type} option for `#{opt}`. :#{type} can be an integer, range, or :all")
      end
    end

    class NoDataError < ArgumentError
      def initialize
        super("Missing :data or :instances option")
      end
    end

    class MultipleDataSourcesError < ArgumentError
      def initialize
        super("Both :data and :instances options cannot be combined, please choose one.")
      end
    end
    
    class SpreadsheetColumnsNotDefinedError < ArgumentError
      def initialize(klass, method_name='spreadsheet_columns')
        super("The instance method `#{method_name}` is not defined on #{klass}")
      end
    end
  end
end
