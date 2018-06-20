def assert_changed(expression, &block)
  if expression.respond_to?(:call)
    e = expression
  else
    e = lambda{ eval(expression, block.binding) }
  end
  old = e.call
  block.call
  assert_not_equal old, e.call
end

def assert_not_changed(expression, &block)
  if expression.respond_to?(:call)
    e = expression
  else
    e = lambda{ eval(expression, block.binding) }
  end
  old = e.call
  block.call
  assert_equal old, e.call
end
