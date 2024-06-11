# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

This script contains various utility functions for converting Roman numerals to integers, 
calculating soil temperature based on the time of the year, determining the number of unknowns 
in a simulation system, and creating coupling matrices for network simulations. 
It also includes functions for printing text in colored formats.
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

import numpy as np
import os
import termcolor
from datetime import datetime 

def romanToInt(s):
    """Converts roman numerals to integer values.

    :param s: Roman numeral
    :type s: str
    :return: Integer representation of s
    :rtype: int
    """
    pass

    roman = {'I':1,'V':5,'X':10,'L':50,'C':100,'D':500,'M':1000,'IV':4,'IX':9,'XL':40,'XC':90,'CD':400,'CM':900}
    i = 0
    num = 0
    while i < len(s):
        if i+1<len(s) and s[i:i+2] in roman:
            num+=roman[s[i:i+2]]
            i+= 2
        else:
            num+=roman[s[i]]
            i += 1
    return num

def soil_temp(time):
    """Calculates soil temperature for a given time of the year.

    :param time: Time for which the soil temperature is to be calculated
    :type time: datetime

    :return: Soil temperature [°C]
    :rtype: float
    """    

    time_year = time.year
    time_year_start = datetime(time_year ,1 ,1 , 0, 0)
    time_year_end = datetime(time_year+1 ,1 ,1 , 0, 0)

    delta_sec = (time-time_year_start).total_seconds()
    delta_sec_total = (time_year_end-time_year_start).total_seconds()

    year_percent = delta_sec/delta_sec_total

    temp_soil = 682.432357637794*year_percent**5\
        - 1583.11270499113*year_percent**4 \
        + 1123.72932688659*year_percent**3 \
        - 227.735852646234*year_percent**2 \
        + 4.79875094132149*year_percent \
        + 3.63636877831834

    return temp_soil

def number_unknowns(line, node, var_sim, var_H2O):
    """
    Determines the number of unknowns for the equation system and sets up start values.

    :param line: Contains information about the pipes
    :type line: line obj.

    :param node: Contains information about the nodes
    :type node: node obj.

    :param var_sim: Various variables concerning the simulation
    :type var_sim: var_sim obj.
    """ 

    ###########################################################################
    # INITIALIZATION ##########################################################
    ###########################################################################

    var_sim.input_solver_hydr_forerun, var_sim.input_solver_hydr_return, var_sim.input_solver_therm \
        = ([] for i in range(3))

    # start values
    # pipe mass flows [kg/s]
    m_dot_start = var_sim.m_dot_start
    # node pressures [Pa]
    p_start = var_sim.p_start
    # pipe temperatures [°C]
    t_start = var_sim.t_start

    ###########################################################################
    # UNKNOWN INTERNAL MASS FLOWS #############################################
    ###########################################################################
    # The mass flows in all pipes are unknown
    for x_line in range(0, line.nbr_matrix.shape[0]):
        var_sim.input_solver_hydr_forerun.append(m_dot_start)
        var_sim.input_solver_hydr_return.append(m_dot_start)

    ###########################################################################
    # UNKNOWN PRESSURES #######################################################
    ###########################################################################
    # The pressures in all nodes except for the pressure reference node are unknown
    for x_node in range(0, node.nbr_matrix.shape[0]-len(node.p_ref)):
        var_sim.input_solver_hydr_forerun.append(p_start)
        var_sim.input_solver_hydr_return.append(p_start)

    ###########################################################################
    # UNKNOWN EXTERNAL MASS FLOWS #############################################
    ###########################################################################
    # The mass flow at the pressure reference node is unknown
    for x_node in range(0, len(node.feed_in)):
        if (node.feed_in[x_node]) and (x_node in node.p_ref):
            m_ext_current = node.V_dot_feed[x_node][var_sim.cntr_time_hyd] * var_H2O.rho / 1000
            var_sim.input_solver_hydr_forerun.append(m_ext_current)
            var_sim.input_solver_hydr_return.append(-m_ext_current)

    ###########################################################################
    # UNKNOWN TEMPERATURES ####################################################
    ###########################################################################
    # The temperatures in all pipes are unknown
    for x_line in range(0, line.nbr_matrix.shape[0]):
        var_sim.input_solver_therm.append(t_start)
    for x_line in range(0, line.nbr_matrix.shape[0]):
        var_sim.input_solver_therm.append(t_start)


def make_matrix_coupl(line, node, var_sim):
    """Creates the n x m coupling matrix for the given network. 
    n is the number of nodes, m is the number of pipes.
    The, n, m-th entry of the matrix is 1 if the n-th node is the start node 
    of the m-th pipe, -1 if it is the end node and 0 otherwise.

    :param line: Contains information about the pipes
    :type line: line obj.

    :param node: Contains information about the nodes
    :type node: node obj.

    :param var_sim: Various variables concerning the simulation
    :type var_sim: var_sim obj.
    """    
    var_sim.matrix_coupl_forerun, var_sim.matrix_coupl_forerun_trans = [], []

    var_sim.matrix_coupl_forerun = np.zeros((node.nbr_matrix.shape[0], line.nbr_matrix.shape[0]))
    var_sim.matrix_coupl_forerun = var_sim.matrix_coupl_forerun.astype(int)

    for x_line in range(0, line.nbr_matrix.shape[0]):
        var_sim.matrix_coupl_forerun[int(line.node_start[x_line]), x_line] = 1
        var_sim.matrix_coupl_forerun[int(line.node_end[x_line]), x_line] = -1

    # TRANSPONIERTE KOPPLUNGSMATRIX
    var_sim.matrix_coupl_forerun_trans = var_sim.matrix_coupl_forerun.transpose()
    
def print_red(text):
    """Prints text in red.

    :param text: Text to be printed
    :type text: str
    """    
    print(termcolor.colored(text, 'red'))


