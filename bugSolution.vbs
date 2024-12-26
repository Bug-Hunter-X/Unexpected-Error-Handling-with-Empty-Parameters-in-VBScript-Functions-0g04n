Function MyFunction(param1, param2)
  If IsNull(param1) Or IsNull(param2) Or (TypeName(param1) = "Variant" And param1 = "") Or (TypeName(param2) = "Variant" And param2 = "") Then
    Err.Raise 9, , "Parameters cannot be empty or null"
  End If
  ' ... rest of the function ...
End Function