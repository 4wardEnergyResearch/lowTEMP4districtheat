# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Provides functionality to solve thermal equations in a network and visualize the results. 
The main function, `solve_network_therm`, calculates thermal behavior in a network of nodes and pipes 
over time, while the `plot_thermal_eqs` function visualizes these calculations using NetworkX and Matplotlib.

Functions:
- solve_network_therm: Solves the thermal equation system defined by the inputs.
- setup_graph: Sets up a graph for visualization using NetworkX.
- draw_graph: Draws the graph with color mapping for temperatures.
- plot_thermal_eqs: Plots the thermal equations and saves the plot as a PNG file.
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
import math
import collections
import matplotlib
import matplotlib.pyplot as plt
import networkx as nx
from datetime import timedelta
import sys
import tkinter as tk
from pandas import DataFrame
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import *
import time
import pylab as pb
import os

matplotlib.use('TkAgg')

###############################################################################
###############################################################################
# FUNCTION TO SOLVE THE THERMAL EQUATION SYSTEM ###############################
###############################################################################
###############################################################################
def solve_network_therm(var_phy, var_H2O, var_sim, var_misc, line, node, plots):
    """Solves the thermal equation system defined by the inputs.

    :param var_phy: Physical variables
    :type var_phy: var_phy obj.

    :param var_H2O: Water variables
    :type var_H2O: var_H2O obj.

    :param var_sim: Simulation variables
    :type var_sim: var_sim obj.

    :param var_misc: Miscellaneous variables
    :type var_misc: var_misc obj.

    :param line: Contains information about the lines
    :type line: line obj.

    :param node: Contains information about the nodes
    :type node: node obj.

    :param plots: Contains plots and corresponding information
    :type plots: plots obj.
    """

    # Number of thermal calculation steps within one hydraulic calc. step
    cntr_therm_max = int(var_sim.delta_time_hyd*60/var_sim.delta_time_therm)    

    ###########################################################################
    # CALCULATION OF THE THERMAL EQUATION SYSTEM IN EVERY TIME STEP ###########
    ###########################################################################
    for cntr_therm_1 in range (0, cntr_therm_max):

        # INITIALIZATION OF THE CALCULATION STEPS
        # Pipes for which the energy balances have already been solved in the current step
        line_calc_forerun = np.zeros(line.nbr_matrix.shape[0])     
        # Nodes for which the energy balances have already been solved in the current step
        node_calc_forerun = np.zeros(node.nbr_matrix.shape[0])
        # Pipes with the necessary data available to solve the energy balances in the current step
        line_calc_now_forerun = np.zeros(line.nbr_matrix.shape[0]) 
        # Nodes with the necessary data available to solve the energy balances in the current step
        node_calc_now_forerun = np.zeros(node.nbr_matrix.shape[0]) 
        # Pipes for which the energy balances have already been solved in the current step
        line_calc_return = np.zeros(line.nbr_matrix.shape[0]) 
        # Nodes for which the energy balances have already been solved in the current step
        node_calc_return = np.zeros(node.nbr_matrix.shape[0])      
        # Pipes with the necessary data available to solve the energy balances in the current step
        line_calc_now_return = np.zeros(line.nbr_matrix.shape[0])  
        # Nodes with the necessary data available to solve the energy balances in the current step
        node_calc_now_return = np.zeros(node.nbr_matrix.shape[0]) 

        # INITIALIZATION OF NODAL TEMPERATURES
        for x_node in range (0, node.nbr_matrix.shape[0]):
            if var_sim.cntr_time_therm_forerun != 0:
                node.t_forerun_trans[x_node] = np.vstack([node.t_forerun_trans[x_node], 0.0])
            else:
                node.t_forerun_trans.append(np.array([0.0]))

        for x_node in range (0, node.nbr_matrix.shape[0]):
            if var_sim.cntr_time_therm_return != 0:
                node.t_return_trans[x_node] = np.vstack([node.t_return_trans[x_node], 0.0])
            else:
                node.t_return_trans.append(np.array([0.0]))

        # FIRST ENERGY BALANCES OF NODES AND PIPES - FLOW
        for x_node in range (0, node.nbr_matrix.shape[0]):
            sum_line_in = collections.Counter(var_sim.matrix_coupl_forerun[x_node])[-1]
            # Find the starting points of the network
            if (node.m_ext_forerun[x_node] >= 0) and (sum_line_in == 0):
                #break
                # Set forerun temperature of feeder to its measured forerun temp.
                node.t_forerun_trans[x_node][var_sim.cntr_time_therm_forerun] = node.temp_flow_feed[x_node][var_sim.cntr_time_hyd]
                # Iterate through lines connected to the node
                for x_line in range (0, line.nbr_matrix.shape[0]):
                    if var_sim.matrix_coupl_forerun[x_node, x_line] == 1:
                        # Calculate the temperature gradients
                        line.dTdt[x_line][1:int(line.n[x_line])] = (line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(line.t_forerun[x_line][0:int(line.n[x_line]-1)]-line.t_forerun[x_line][1:int(line.n[x_line])])-line.htc[x_line]*(line.t_forerun[x_line][1:int(line.n[x_line])]-var_sim.temp_soil[-1])*math.pi*line.dia[x_line]*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*line.dia[x_line]**2/4)
                        line.dTdt[x_line][0] = (line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(node.temp_flow_feed[x_node][var_sim.cntr_time_hyd]-line.t_forerun[x_line][0])-line.htc[x_line]*(line.t_forerun[x_line][0]-var_sim.temp_soil[-1])*math.pi*line.dia[x_line]*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*line.dia[x_line]**2/4)
                        line.t_forerun[x_line] = line.t_forerun[x_line]+line.dTdt[x_line]*var_sim.delta_time_therm
                        if var_sim.cntr_time_therm_forerun != 0:
                            line.t_forerun_trans[x_line] = np.vstack([line.t_forerun_trans[x_line], line.t_forerun[x_line]])
                        else:
                            line.t_forerun_trans[x_line] = line.t_forerun[x_line][np.newaxis,:]
                        line_calc_forerun[x_line] = 1
                        node_calc_forerun[x_node] = 1

###############################################################################
# FLOW ########################################################################
###############################################################################
        calc_therm = 1
        while calc_therm == 1:

            # FIND NODES WHICH CAN BE SOLVED AT CURRENT SIMULATION STEP
            for x_node in range (0, node.nbr_matrix.shape[0]):
                if node_calc_forerun[x_node] == 0:
                    sum_t_int_out_unkn = 0
                    sum_t_int_out_known = 0
                    for x_line in range (0, line.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_forerun[x_node, x_line] == -1 and line_calc_forerun[x_line] == 0:
                            # Number of unknown outgoing internal mass flows
                            sum_t_int_out_unkn = sum_t_int_out_unkn+1
                    for x_line in range (0, line.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_forerun[x_node, x_line] == -1 and line_calc_forerun[x_line] == 1:
                            # Number of known outgoing internal mass flows
                            sum_t_int_out_known = sum_t_int_out_known+1
                    if sum_t_int_out_unkn == 0 and sum_t_int_out_known > 0: # If all outgoing internal mass flows are known, the node can be solved
                            node_calc_now_forerun[x_node] = 1

            # ENERGY BALANCES OF THE NODES WHICH CAN BE SOLVED FOR THE CURRENT SIMULATION STEP
            for x_node in range (0, node.nbr_matrix.shape[0]):
                if node_calc_now_forerun[x_node] == 1:
                    sum_line_out = collections.Counter(var_sim.matrix_coupl_forerun[x_node])[1]
                    sum_line_in = collections.Counter(var_sim.matrix_coupl_forerun[x_node])[-1]

                    ###########################################################
                    ###########################################################
                    # NODES WITHOUT EXTERNAL MASS FLOWS #######################
                    ###########################################################
                    ###########################################################
                    if node.m_ext_forerun_trans[x_node][var_sim.cntr_time_hyd] == 0:
    
                        #######################################################
                        # NODES WITHOUT EXTERNAL MASS FLOW AND WITH ONE #######
                        # INCOMING INTERNAL MASS FLOW #########################
                        #######################################################
                        if sum_line_in == 1:
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_forerun[x_node, x_line] == -1:
                                    x_line_ref = x_line
                                    break

                            node.t_forerun_trans[x_node][var_sim.cntr_time_therm_forerun] = line.t_forerun[x_line_ref][-1]
                            node_calc_forerun[x_node] = 1

                        #######################################################
                        # NODES WITHOUT EXTERNAL MASS FLOWS AND WITH MORE #####
                        # THAN ONE INCOMING INTERNAL MASS FLOW ################
                        #######################################################
                        else:
                            sum_1 = 0
                            sum_2 = 0
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_forerun[x_node, x_line] == -1:
                                    add_1 = line.t_forerun[x_line][-1]*line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]
                                    add_2 = line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]
                                    sum_1 = sum_1+add_1
                                    sum_2 = sum_2+add_2
                            node.t_forerun_trans[x_node][var_sim.cntr_time_therm_forerun] = sum_1/sum_2
                            node_calc_forerun[x_node] = 1

                    ###########################################################
                    ###########################################################
                    # NODES WITH AN EXTERNAL CONSUMER #########################
                    ###########################################################
                    ###########################################################
                    # Here, we can ignore the external consumer, as it does not change the temperature
                    elif node.m_ext_forerun_trans[x_node][var_sim.cntr_time_hyd] < 0:

                        #######################################################
                        # NODE WITH EXTERNAL CONSUMER AND ONE INCOMING ########
                        # INTERNAL MASS FLOW ##################################
                        #######################################################
                        if sum_line_in == 1:
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_forerun[x_node, x_line] == -1:
                                    x_line_ref = x_line
                                    break

                            node.t_forerun_trans[x_node][var_sim.cntr_time_therm_forerun] = line.t_forerun[x_line_ref][-1]
                            node_calc_forerun[x_node] = 1

                        #######################################################
                        # NODES WITH AN EXTERNAL CONSUMER AND MORE THAN ONE ###
                        # INCOMING INTERNAL MASS FLOW #########################
                        #######################################################
                        else:
                            sum_1 = 0
                            sum_2 = 0
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_forerun[x_node, x_line] == -1:
                                    add_1 = line.t_forerun[x_line][-1]*line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]
                                    add_2 = line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]
                                    sum_1 = sum_1+add_1
                                    sum_2 = sum_2+add_2
                            node.t_forerun_trans[x_node][var_sim.cntr_time_therm_forerun] = sum_1/sum_2
                            node_calc_forerun[x_node] = 1

                    ###########################################################
                    ###########################################################
                    # NODES WITH AN EXTERNAL FEEDER ###########################
                    ###########################################################
                    ###########################################################              
                    else:

                        #######################################################
                        # NODE WITH AN EXTERNAL FEEDER AND ONE OR MORE ########
                        # INCOMING INTERNAL MASS FLOWS ########################
                        #######################################################
                        if sum_line_in > 0:
                            # Set external mass flow of feeder
                            if (node.feed_in[x_node]) and (x_node in node.p_ref):
                                m_ext_current = node.V_dot_feed[x_node][var_sim.cntr_time_hyd] * var_H2O.rho / 1000
                            else:
                                m_ext_current = node.m_ext_forerun[x_node][var_sim.cntr_time_hyd]

                            sum_1 = node.temp_flow_feed[x_node][var_sim.cntr_time_hyd]*m_ext_current
                            sum_2 = m_ext_current
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_forerun[x_node, x_line] == -1:
                                    add_1 = line.t_forerun[x_line][-1]*line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]
                                    add_2 = line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]
                                    sum_1 = sum_1+add_1
                                    sum_2 = sum_2+add_2
                            node.t_forerun_trans[x_node][var_sim.cntr_time_therm_forerun] = \
                                (sum_1) / \
                                (sum_2)
                            node_calc_forerun[x_node] = 1

            # END OF NODE CALCULATION
            node_calc_now_forerun = np.zeros(node.nbr_matrix.shape[0])

            # DETERMINE PIPES WHICH CAN BE SOLVED AT THE CURRENT SIMULATION STEP
            for x_line in range (0, line.nbr_matrix.shape[0]):
                if line_calc_forerun[x_line] == 0:
                    for x_node in range (0, node.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_forerun[x_node, x_line] == 1 and node_calc_forerun[x_node] == 1:
                            line_calc_now_forerun[x_line] = 1
                            break

            # TRANSIENT THERMAL CALCULATIONS OF THE PIPES WHICH CAN BE PERFORMED AT THE CURRENT SIMULATION STEP
            for x_line in range (0, line.nbr_matrix.shape[0]):
                if line_calc_now_forerun[x_line] == 1:
                    for x_node in range (0, node.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_forerun[x_node, x_line] == 1:
                            x_node_ref = x_node
                            break
                    line.dTdt[x_line][0] = (line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(node.t_forerun_trans[x_node_ref][var_sim.cntr_time_therm_forerun]-line.t_forerun[x_line][0])-line.htc[x_line]*(line.t_forerun[x_line][0]-var_sim.temp_soil[-1])*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*(line.dia[x_line]**2)/4)
                    line.dTdt[x_line][1:int(line.n[x_line])] = (line.m_int_forerun_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(line.t_forerun[x_line][0:int(line.n[x_line]-1)]-line.t_forerun[x_line][1:int(line.n[x_line])])-line.htc[x_line]*(line.t_forerun[x_line][1:int(line.n[x_line])]-var_sim.temp_soil[-1])*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*(line.dia[x_line]**2)/4)
                    
                    line.t_forerun[x_line] = line.t_forerun[x_line]+line.dTdt[x_line]*var_sim.delta_time_therm
                    if var_sim.cntr_time_therm_forerun != 0:
                        line.t_forerun_trans[x_line] = np.vstack([line.t_forerun_trans[x_line], line.t_forerun[x_line]])
                    else:
                        line.t_forerun_trans[x_line] = line.t_forerun[x_line][np.newaxis,:]
                    line_calc_forerun[x_line] = 1

            # END OF PIPE CALCULATION
            line_calc_now_forerun = np.zeros(line.nbr_matrix.shape[0])
            
            if sum(line_calc_forerun) == len(line_calc_forerun) and sum(node_calc_forerun) == len(node_calc_forerun):
                calc_therm = 0
                var_sim.cntr_time_therm_forerun += 1

###############################################################################
# RETURN ######################################################################
###############################################################################

        # FIRST ENERGY BALANCES OF NODES AND PIPES - RETURN
        for x_node in range (0, node.nbr_matrix.shape[0]):
            sum_line_in = collections.Counter(var_sim.matrix_coupl_return[x_node])[-1]
            # Find points to start the calculation
            if (node.m_ext_return[x_node] >= 0) and (sum_line_in == 0):
                if node.m_ext_return[x_node] == 0:
                    # If there is no flow at the current time step:                    
                    if var_sim.cntr_time_hyd == 0:
                        # If this is the first time step, the temperature of the return is set to the temperature of the forerun
                        node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = node.t_forerun_trans[x_node][var_sim.cntr_time_therm_return]
                    else:
                        # Set the temperature of the return to the temperature of the forerun minus the last known temperature difference
                        deltaT_last = node.t_forerun_trans[x_node][var_sim.cntr_time_therm_return - 1] - node.t_return_trans[x_node][var_sim.cntr_time_therm_return - 1]
                        node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = node.t_forerun_trans[x_node][var_sim.cntr_time_therm_return] - deltaT_last
                else:
                    # If it's a consumer, everything is fine.
                    if node.feed_in[x_node] == None:
                        V_current = node.m_ext_return[x_node]/var_H2O.rho * 1000
                        Q_current = node.Q_dot_sim[x_node][var_sim.cntr_time_hyd]
                        t_current = node.t_forerun_trans[x_node][var_sim.cntr_time_therm_return] - Q_current*1000/(V_current/1000*var_H2O.rho*var_H2O.c_p)
                    # If it's a feeder, this point should not be reached. If it is reached regardlessly, this indicates that the network topology is not correct.
                    else:
                        raise ValueError(f"Umkehr der Strömungsrichtung bei Einspeiser {node.feed_in[x_node]} festgestellt!")

                    if t_current < var_sim.temp_soil[-1]:
                        node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = var_sim.temp_soil[-1]
                    else:
                        node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = t_current

                # Calculate outgoing lines
                for x_line in range (0, line.nbr_matrix.shape[0]):
                    if var_sim.matrix_coupl_return[x_node, x_line] == 1:
                        line.dTdt[x_line][0] = (line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(node.t_return_trans[x_node][var_sim.cntr_time_therm_return]-line.t_return[x_line][0])-line.htc[x_line]*(line.t_return[x_line][0]-var_sim.temp_soil[-1])*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*(line.dia[x_line]**2)/4)
                        line.dTdt[x_line][1:int(line.n[x_line])] = (line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(line.t_return[x_line][0:int(line.n[x_line]-1)]-line.t_return[x_line][1:int(line.n[x_line])])-line.htc[x_line]*(line.t_return[x_line][1:int(line.n[x_line])]-var_sim.temp_soil[-1])*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*(line.dia[x_line]**2)/4)
                        line.t_return[x_line] = line.t_return[x_line]+line.dTdt[x_line]*var_sim.delta_time_therm
                        if var_sim.cntr_time_therm_return != 0:
                            line.t_return_trans[x_line] = np.vstack([line.t_return_trans[x_line], line.t_return[x_line]])
                        else:
                            line.t_return_trans[x_line] = line.t_return[x_line][np.newaxis,:]
                        line_calc_return[x_line] = 1
                        node_calc_return[x_node] = 1

        calc_therm = 1
        while calc_therm == 1:
            # FIND NODES WHICH CAN BE SOLVED AT CURRENT SIMULATION STEP
            for x_node in range (0, node.nbr_matrix.shape[0]):
                if node_calc_return[x_node] == 0:
                    sum_t_int_out_unkn = 0
                    sum_t_int_out_known = 0
                    for x_line in range (0, line.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_return[x_node, x_line] == -1 and line_calc_return[x_line] == 0:
                            sum_t_int_out_unkn = sum_t_int_out_unkn+1
                    for x_line in range (0, line.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_return[x_node, x_line] == -1 and line_calc_return[x_line] == 1:
                            sum_t_int_out_known = sum_t_int_out_known+1
                    if sum_t_int_out_unkn == 0 and sum_t_int_out_known > 0:
                            node_calc_now_return[x_node] = 1

            # ENERGY BALANCES OF THE NODES WHICH CAN BE SOLVED FOR THE CURRENT SIMULATION STEP
            for x_node in range (0, node.nbr_matrix.shape[0]):
                if node_calc_now_return[x_node] == 1:
                    sum_line_out = collections.Counter(var_sim.matrix_coupl_return[x_node])[1]
                    sum_line_in = collections.Counter(var_sim.matrix_coupl_return[x_node])[-1]

                    ###########################################################
                    ###########################################################
                    # NODES WITHOUT EXTERNAL MASS FLOWS #######################
                    ###########################################################
                    ###########################################################
                    if node.m_ext_return_trans[x_node][var_sim.cntr_time_hyd] == 0:
    
                        #######################################################
                        # NODES WITHOUT EXTERNAL MASS FLOW AND WITH ONE #######
                        # INCOMING INTERNAL MASS FLOW #########################
                        #######################################################
                        if sum_line_in == 1:
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_return[x_node, x_line] == -1:
                                    x_line_ref = x_line
                                    break

                            node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = line.t_return[x_line_ref][-1]
                            node_calc_return[x_node] = 1

                        #######################################################
                        # NODES WITHOUT EXTERNAL MASS FLOWS AND WITH MORE #####
                        # THAN ONE INCOMING INTERNAL MASS FLOW ################
                        #######################################################
                        else:
                            sum_1 = 0
                            sum_2 = 0
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_return[x_node, x_line] == -1:
                                    add_1 = line.t_return[x_line][-1]*line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]
                                    add_2 = line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]
                                    sum_1 = sum_1+add_1
                                    sum_2 = sum_2+add_2
                            if sum_2 != 0:
                                node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = sum_1/sum_2
                            else:
                                node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = node.t_return_trans[x_node][var_sim.cntr_time_therm_return-1]
                            node_calc_return[x_node] = 1

                    ###########################################################
                    ###########################################################
                    # NODES WITH AN EXTERNAL CONSUMER #########################
                    ###########################################################
                    ###########################################################
                    elif (node.m_ext_return_trans[x_node][var_sim.cntr_time_hyd] > 0) and (node.feed_in[x_node] == None):
                        # Calculate the temperature of the external return stream
                        temp_ext_return = node.t_forerun_trans[x_node][var_sim.cntr_time_therm_return]-node.Q_dot_sim[x_node][var_sim.cntr_time_hyd]*1000/ \
                                            (node.V_dot_sim[x_node][var_sim.cntr_time_hyd]/1000*var_H2O.rho*var_H2O.c_p)
                        if  temp_ext_return < var_sim.temp_soil[-1]:
                            node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = var_sim.temp_soil[-1]
                        else:
                            node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = temp_ext_return

                        #######################################################
                        # NODES WITH AN EXTERNAL CONSUMER AND AT LEAST ONE ####
                        # INCOMING INTERNAL MASS FLOW #########################
                        #######################################################
                        # We need no case differentiation here. This works the same for one or more incoming internal mass flows.
                        # Add external return stream (This can be done in the sum terms at initialization)
                        sum_1 = temp_ext_return * node.m_ext_return_trans[x_node][var_sim.cntr_time_hyd]
                        sum_2 = node.m_ext_return_trans[x_node][var_sim.cntr_time_hyd]
                        for x_line in range (0, line.nbr_matrix.shape[0]):
                            if var_sim.matrix_coupl_return[x_node, x_line] == -1:
                                add_1 = line.t_return[x_line][-1]*line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]
                                add_2 = line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]
                                sum_1 = sum_1+add_1
                                sum_2 = sum_2+add_2
                        if sum_2 != 0:
                            node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = sum_1/sum_2
                        else:
                            node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = node.t_return_trans[x_node][var_sim.cntr_time_therm_return-1]
                        node_calc_return[x_node] = 1

                    ###########################################################
                    ###########################################################
                    # NODES WITH AN EXTERNAL FEEDER ###########################
                    ###########################################################
                    ###########################################################

                    else:
                        
                        #######################################################
                        # NODE WITH AN EXTERNAL FEEDER AND ONE OR MORE ########
                        # INCOMING INTERNAL MASS FLOWS ########################
                        #######################################################
                        # The feeder can be ignored since it does not change the temperature.
                        if sum_line_in > 0:
                            sum_1 = 0
                            sum_2 = 0
                            for x_line in range (0, line.nbr_matrix.shape[0]):
                                if var_sim.matrix_coupl_return[x_node, x_line] == -1:
                                    add_1 = line.t_return[x_line][-1]*line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]
                                    add_2 = line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]
                                    sum_1 = sum_1+add_1
                                    sum_2 = sum_2+add_2
                            node.t_return_trans[x_node][var_sim.cntr_time_therm_return] = sum_1/sum_2
                            node_calc_return[x_node] = 1

            # END OF NODE CALCULATION
            node_calc_now_return = np.zeros(node.nbr_matrix.shape[0])

            # DETERMINE PIPES WHICH CAN BE SOLVED AT THE CURRENT SIMULATION STEP
            for x_line in range (0, line.nbr_matrix.shape[0]):
                if line_calc_return[x_line] == 0:
                    for x_node in range (0, node.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_return[x_node, x_line] == 1 and node_calc_return[x_node] == 1:
                            line_calc_now_return[x_line] = 1
                            break

            # TRANSIENT THERMAL CALCULATIONS OF THE PIPES WHICH CAN BE PERFORMED AT THE CURRENT SIMULATION STEP
            for x_line in range (0, line.nbr_matrix.shape[0]):
                if line_calc_now_return[x_line] == 1:
                    for x_node in range (0, node.nbr_matrix.shape[0]):
                        if var_sim.matrix_coupl_return[x_node, x_line] == 1:
                            x_node_ref = x_node
                            break
                    line.dTdt[x_line][1:int(line.n[x_line])] = (line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(line.t_return[x_line][0:int(line.n[x_line]-1)]-line.t_return[x_line][1:int(line.n[x_line])])-line.htc[x_line]*(line.t_return[x_line][1:int(line.n[x_line])]-var_sim.temp_soil[-1])*math.pi*line.dia[x_line]*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*line.dia[x_line]**2/4)
                    line.dTdt[x_line][0] = (line.m_int_return_trans[x_line][var_sim.cntr_time_hyd]*var_H2O.c_p*(node.t_return_trans[x_node_ref][var_sim.cntr_time_therm_return]-line.t_return[x_line][0])-line.htc[x_line]*(line.t_return[x_line][0]-var_sim.temp_soil[-1])*math.pi*line.dia[x_line]*line.dx[x_line])/(var_H2O.rho*var_H2O.c_p*line.dx[x_line]*math.pi*line.dia[x_line]**2/4)
                    line.t_return[x_line] = line.t_return[x_line]+line.dTdt[x_line]*var_sim.delta_time_therm
                    if var_sim.cntr_time_therm_return != 0:
                        line.t_return_trans[x_line] = np.vstack([line.t_return_trans[x_line], line.t_return[x_line]])
                    else:
                        line.t_return_trans[x_line] = line.t_return[x_line][np.newaxis,:]
                    line_calc_return[x_line] = 1

            # END OF PIPE CALCULATION
            line_calc_now_return = np.zeros(line.nbr_matrix.shape[0])
            
            if sum(line_calc_return) == len(line_calc_return) and sum(node_calc_return) == len(node_calc_return):
                calc_therm = 0
                var_sim.cntr_time_therm_return += 1

        # Every nth time step, the thermal calculations are plotted
        if (var_sim.cntr_time_therm_forerun - 1) % plots.thermal_update_time_steps == 0:
            plot_thermal_eqs(var_sim, line, node, plots)

def setup_graph(line_data, node_data, var_sim, forerun):
    """
    Setup the graph for the thermal calculations.

    :param line_data: Contains the line data
    :type line_data: line_data obj.
    :param node_data: Contains the node data
    :type node_data: node_data obj.
    :param var_sim: Contains the simulation data
    :type var_sim: var_sim obj.
    :param forerun: Boolean indicating if the graph is for the forerun or return simulation
    :type forerun: bool
    :return: Graph object and label dictionary
    :rtype: nx.Graph, dict
    """

    graph = nx.Graph()
    label_dict = {}
    color_list = []
    cntr_node = 1

    matrix = var_sim.matrix_coupl_forerun if forerun else var_sim.matrix_coupl_return
    t_trans = line_data.t_forerun_trans if forerun else line_data.t_return_trans

    for x_line in range(0, line_data.nbr_matrix.shape[0]):
        # Get start and end nodes from coupling matrix
        # Get column corresponding to the current line
        column = matrix[:, x_line]
        # Get the start and end nodes of the line
        start_node = np.where(column == 1)[0][0]
        end_node = np.where(column == -1)[0][0]

        line_seg = len(t_trans[x_line][0])
        x_coords = np.linspace(node_data.x_coord[int(start_node)],
                               node_data.x_coord[int(end_node)], line_seg + 1)
        y_coords = np.linspace(node_data.y_coord[int(start_node)],
                               node_data.y_coord[int(end_node)], line_seg + 1)
        for x_seg in range(line_seg + 1):
            graph.add_node(cntr_node, pos=(x_coords[x_seg], y_coords[x_seg]))
            label_dict[cntr_node] = 1
            if x_seg > 0:
                graph.add_edge(cntr_node - 1, cntr_node, weight=3 * line_data.dia[x_line] / max(line_data.dia))
            cntr_node += 1
    return graph, label_dict

def draw_graph(graph, var_line_data, title, subplot_pos, current_sim_time, cmap, edge_colors, vmin, vmax, fig):
    """
    Draw the graph for the thermal calculations.

    :param graph: Graph object
    :type graph: nx.Graph
    :param var_line_data: Contains the line data
    :type var_line_data: line_data obj.
    :param title: Title of the plot
    :type title: str
    :param subplot_pos: Position of the subplot
    :type subplot_pos: int
    :param current_sim_time: Current simulation time
    :type current_sim_time: datetime
    :param cmap: Colormap for the edge colors
    :type cmap: plt.cm
    :param edge_colors: Colors for the edges
    :type edge_colors: np.array
    :param vmin: Minimum value for the color mapping
    :type vmin: float
    :param vmax: Maximum value for the color mapping
    :type vmax: float
    :param fig: Figure object
    :type fig: plt.figure
    """

    ax = fig.add_subplot(subplot_pos)  # Create subplot in the specified position
    pos = nx.get_node_attributes(graph, 'pos')
    weights = [graph[u][v]['weight'] for u, v in graph.edges()]

    # Draw only the edges with color mapping.
    edge_viz = nx.draw_networkx_edges(
        graph, pos, edge_color=edge_colors, edge_cmap=cmap, width=weights, ax=ax
    )
    plt.title(f"{title} - {current_sim_time.strftime('%Y-%m-%d %H:%M:%S')}")

    # Add colorbar for edge color mapping.
    sm = plt.cm.ScalarMappable(cmap=cmap, norm=plt.Normalize(vmin=vmin, vmax=vmax))
    sm.set_array([])
    fig.colorbar(sm, ax=ax, label=f"{title} temperature [°C]")  # Add colorbar to the specific subplot axis


def plot_thermal_eqs(var_sim, line, node, plots):
    """
    Plot the thermal equations.
    
    :param var_sim: Contains the simulation data
    :type var_sim: var_sim obj.
    :param line: Contains the line data
    :type line: line_data obj.
    :param node: Contains the node data
    :type node: node_data obj.
    :param plots: Contains the plot settings
    :type plots: plots obj.
    """
    if plots.show_plot == "yes":
        if not plots.fig:
            plots.fig = plt.figure(1, figsize=(21, 9))
        
        # Change active figure to fig
        plt.figure(1)

        plots.fig.clear()
        current_sim_time = var_sim.time_sim_start + timedelta(seconds=(var_sim.cntr_time_therm_forerun - 1) * var_sim.delta_time_therm)
        cmap = plt.cm.plasma

        # Forerun
        G_forerun, _ = setup_graph(line, node, var_sim, forerun=True)
        colors_forerun = np.concatenate([line.t_forerun_trans[x_line][-1] for x_line in range(line.nbr_matrix.shape[0])])
        vmin_forerun, vmax_forerun = colors_forerun.min(), colors_forerun.max()
        draw_graph(G_forerun, line, "Forerun", 121, current_sim_time, cmap, colors_forerun, vmin_forerun, vmax_forerun, plots.fig)
        
        # Return
        G_return, _ = setup_graph(line, node, var_sim, forerun=False)
        colors_return = np.concatenate([line.t_return_trans[x_line][-1] for x_line in range(line.nbr_matrix.shape[0])])
        vmin_return, vmax_return = colors_return.min(), colors_return.max()
        draw_graph(G_return, line, "Return", 122, current_sim_time, cmap, colors_return, vmin_return, vmax_return, plots.fig)
        
        plt.show()
        plt.pause(0.1)  # Pause to allow GUI to update

        # Save the plot
        filename = f"{plots.topology_file_name}_THERM_{current_sim_time.strftime('%Y%m%d_%H%M%S')}.png"
        path = os.path.join(plots.output_dir, filename)
        plt.savefig(path, dpi=200)
