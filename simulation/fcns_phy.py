# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Functions used for modelling physical quantities.

Functions:
- H2O_density: Calculates the density of water at a given temperature.
"""
"""
Copyright (C) 2024  4wardEnergy Research GmbH

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

def H2O_density (temp):
    """returns the density of water in kg/m³ at 1013 mbar as a function of temperature.

    :param temp: Temperature [°C]
    :type temp: float

    :return: Density of water [kg/m³]
    :rtype: float
    """    

    rho = 999.972-7*10**(-3)*(temp-4)
    # http://109.205.171.104/~thomas/eth/3_semester/hydrosphaere_WS_2004_2005/unterlagen/skript_1.pdf

    return (rho)