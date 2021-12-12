from CoolProp.CoolProp import PropsSI

print(PropsSI("H", 'T', 870, 'P', 28001325, 'IF97::Water'))

print(PropsSI("P", 'T', 373, 'Q', 0, 'IF97::Water'))

print(PropsSI("T", 'P', 101325, 'Q', 0, 'IF97::Water')-273.15)

# Critical temperature/Pressure for Water
print(PropsSI('Tcrit', 'T', 0, 'P', 0, 'IF97::Water'))
print(PropsSI('Pcrit', 'T', 0, 'P', 0, 'IF97::Water'))

heat = (PropsSI("H", 'P', 11080000, 'T', 310.4 + 273.1, 'IF97::Water') - PropsSI("H", 'P', 11080000, 'T', 297.6 + 273.1,
                                                                                'IF97::Water')) * 280504 / 1e9

heat2 = (PropsSI("H", 'P', 11010000, 'T', 310.4 + 273.1, 'IF97::Water') - PropsSI("H", 'P', 11010000, 'T', 295.32 + 273.1,
                                                                                'IF97::Water')) * 286443 / 1e9


print(heat, heat2)

print((2.61404656177333000000 * 79.56 + 90.05644800378120000000)*23.2 )