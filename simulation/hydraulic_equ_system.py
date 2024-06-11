# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Functions to solve and visualize the hydraulic equation system for a network.

Functions:
- solve_network_hydr: Solves the hydraulic equation system defined by the inputs.
- equ_network_return: Creates the hydraulic equation system for the return flow.
- equ_network_forerun: Creates the hydraulic equation system for the forerun flow.
- setup_or_clear_subplots: Sets up subplots for visualization.
- draw_graph: Draws the graph of the network with pressures as node colors.
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

import scipy.optimize as opt
import numpy as np
import math
import matplotlib
import matplotlib.pyplot as plt
import networkx as nx
from datetime import timedelta
import pylab as pb
import os

matplotlib.use('TkAgg')

plt.ion()
# Ignore matplotlib warning
import warnings
warnings.filterwarnings("ignore")



###############################################################################
###############################################################################
# FUNCTION TO SOLVE THE HYDRAULIC EQUATION SYSTEM (CONTINUITY & HYDRAULICS) ###
###############################################################################
###############################################################################
def solve_network_hydr(var_phy, var_H2O, var_sim, var_misc, line, node, plots):
    """Solves the hydraulic equation system defined by the inputs.

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

    # Initialization of the coupling matrices for the return
    var_sim.matrix_coupl_return, var_sim.matrix_coupl_return_trans = [],[]

###############################################################################
# FLOW ########################################################################
###############################################################################

    # ITERATION LOOP
    cntr_wrong_direction = 1
    while cntr_wrong_direction > 0:
        
        # SOLVING THE HYDRAULIC EQUATION SYSTEM FOR THE FLOW
        var_sim.unkn_system_hydr_forerun = opt.fsolve(equ_network_forerun, var_sim.input_solver_hydr_forerun, args = (line, node, var_sim, var_H2O, var_phy))

        # CHECK FLOW DIRECTIONS
        cntr_wrong_direction = 0
        for x_line in range(0, line.nbr_matrix.shape[0]):
            if line.m_int_forerun[x_line] < -1e-10:
                cntr_wrong_direction = cntr_wrong_direction+1
                print(f"Forerun: The flow direction of pipe {int(line.nbr_orig[x_line])} is being corrected...")
                for x_node in range(0, node.nbr_matrix.shape[0]):
                    if var_sim.matrix_coupl_forerun[x_node, x_line] == 1:
                        var_sim.matrix_coupl_forerun[x_node, x_line] = -1
                    elif var_sim.matrix_coupl_forerun[x_node, x_line] == -1:
                        var_sim.matrix_coupl_forerun[x_node, x_line] = 1

        # CALC. TRANSPOSED COUPLING MATRIX ANEW
        var_sim.matrix_coupl_forerun_trans = var_sim.matrix_coupl_forerun.transpose()

###############################################################################
# RETURN ######################################################################
###############################################################################

    # COUPLING MATRIX FOR RETURN
    var_sim.matrix_coupl_return = var_sim.matrix_coupl_forerun*(-1)
    var_sim.matrix_coupl_return_trans = var_sim.matrix_coupl_return.transpose()

    # ITERATION LOOP
    cntr_wrong_direction = 1
    while cntr_wrong_direction > 0:
        
        # SOLVING THE HYDRAULIC EQUATION SYSTEM FOR THE RETURN
        var_sim.unkn_system_hydr_return = opt.fsolve(equ_network_return, var_sim.input_solver_hydr_return, args = (line, node, var_sim, var_H2O, var_phy))

        # CHECK FLOW DIRECTIONS
        cntr_wrong_direction = 0
        for x_line in range(0, line.nbr_matrix.shape[0]):
            if line.m_int_return[x_line] < -1e-10:
                cntr_wrong_direction = cntr_wrong_direction+1
                print(f"Return: The flow direction of pipe {int(line.nbr_orig[x_line])} is being corrected...")
                for x_node in range(0, node.nbr_matrix.shape[0]):
                    if var_sim.matrix_coupl_return[x_node, x_line] == 1: 
                        var_sim.matrix_coupl_return[x_node, x_line] = -1
                    elif var_sim.matrix_coupl_return[x_node, x_line] == -1:
                        var_sim.matrix_coupl_return[x_node, x_line] = 1

        # CALC. TRANSPOSED COUPLING MATRIX ANEW
        var_sim.matrix_coupl_return_trans = var_sim.matrix_coupl_return.transpose()

    # WRITE MASS FLOWS TO ARRAYS
    for x_node in range (0, node.nbr_matrix.shape[0]):
        if var_sim.cntr_time_hyd == 0:
            node.m_ext_forerun_trans.append(np.array([node.m_ext_forerun[x_node]]))
            node.m_ext_return_trans.append(np.array([node.m_ext_return[x_node]]))
        else:
            node.m_ext_forerun_trans[x_node] = np.vstack([node.m_ext_forerun_trans[x_node], node.m_ext_forerun[x_node]])
            node.m_ext_return_trans[x_node] = np.vstack([node.m_ext_return_trans[x_node], node.m_ext_return[x_node]])
    for x_line in range (0, line.nbr_matrix.shape[0]):
        if var_sim.cntr_time_hyd == 0:
            line.m_int_forerun_trans.append(np.array([line.m_int_forerun[x_line]]))
            line.m_int_return_trans.append(np.array([line.m_int_return[x_line]]))
        else:
            line.m_int_forerun_trans[x_line] = np.vstack([line.m_int_forerun_trans[x_line], line.m_int_forerun[x_line]])
            line.m_int_return_trans[x_line] = np.vstack([line.m_int_return_trans[x_line], line.m_int_return[x_line]])

    # WRITE PRESSURES TO ARRAYS
    p_forerun_add, p_return_add = [], []
    for x_node in range(0, node.nbr_matrix.shape[0]):
        offset = node.p_offset[x_node]
        p_forerun_add = np.append(p_forerun_add, np.array([node.p_forerun[x_node] + offset]), axis = 0)
        p_return_add = np.append(p_return_add, np.array([node.p_return[x_node] + offset]), axis = 0)
    node.p_forerun_trans.append(p_forerun_add)
    node.p_return_trans.append(p_return_add)

###############################################################################
###############################################################################
# VISUALIZATION ###############################################################
###############################################################################
###############################################################################

    if plots.show_plot == "yes":

        if not hasattr(plots, 'colorbars'):
            plots.colorbars = []

        ax1, ax2, cbar_ax1, cbar_ax2 = setup_or_clear_subplots(plots)
        
        # Draw the forerun graph
        scalar_map_1 = draw_graph(ax1, var_sim.matrix_coupl_forerun, node.p_forerun_trans[var_sim.cntr_time_hyd], node, line, plots)

        # Draw the return graph
        scalar_map_2 = draw_graph(ax2, var_sim.matrix_coupl_return, node.p_return_trans[var_sim.cntr_time_hyd], node, line, plots)


        cbar1 = plt.colorbar(scalar_map_1, cax=cbar_ax1, label='Forerun Pressure [kPa]')
        plots.colorbars.append(cbar1)  # Keep track of the colorbar

        cbar2 = plt.colorbar(scalar_map_2, cax=cbar_ax2, label='Return Pressure [kPa]')
        plots.colorbars.append(cbar2)  # Keep track of the second colorbar    

        current_sim_time = var_sim.time_sim_start + timedelta(seconds=(var_sim.delta_time_hyd * 60 * var_sim.cntr_time_hyd))

        ax1.set_title(f'Forerun - {current_sim_time.strftime("%Y-%m-%d %H:%M:%S")}')
        ax2.set_title(f'Return - {current_sim_time.strftime("%Y-%m-%d %H:%M:%S")}')

        # Adjust layout
        plt.tight_layout()

        # Show/update the graph without blocking
        plt.draw()
        plt.pause(0.1)  # Allows GUI backend to update the plot

        # Save the plot
        filename = f"{plots.topology_file_name}_HYD_{current_sim_time.strftime('%Y%m%d_%H%M%S')}.png"
        path = os.path.join(plots.output_dir, filename)
        plt.savefig(path, dpi=200)


                
###############################################################################
###############################################################################
# FUNCTION TO CREATE THE HYDRAULIC EQUATION SYSTEM FOR THE RETURN #############
###############################################################################
###############################################################################
def equ_network_return(start_variables_return, line, node, var_sim, var_H2O, var_phy):
    """Creates the hydraulic equation system corresponding to the inputs 
    for the forerun: Continuity equations for each node and pressure equations 
    for each line.

    :param start_variables_return: Start variables for the equation system
    :type start_variables_return: numpy.ndarray

    :param line: Contains information about the lines
    :type line: line obj.

    :param node: Contains information about the nodes
    :type node: node obj.

    :param var_sim: Simulation variables
    :type var_sim: var_sim obj.

    :param var_H2O: Water variables
    :type var_H2O: var_H2O obj.

    :param var_phy: Physical variables
    :type var_phy: var_phy obj.

    :return: Hydraulic equation system
    :rtype: numpy.ndarray
    """    
    # INITIALIZE EQUATION SYSTEMS
    equ_system_hydr_return, equ_system_hydr_return_string = [], []

    # START VALUES FOR UNKNOWNS
    for x_line in range(0, line.nbr_matrix.shape[0]):
        line.m_int_return[x_line] = start_variables_return[x_line]
    cntr_var = x_line+1
    for x_node in range(0, node.nbr_matrix.shape[0]):
        if node.p_ref_return[x_node] == None:
            node.p_return[x_node] = start_variables_return[cntr_var]
            cntr_var = cntr_var+1
    for x_node in range(0, node.m_ext_return.shape[0]):
        if node.m_ext_return_check[x_node] == None:
            node.m_ext_return[x_node] = start_variables_return[cntr_var]
            cntr_var = cntr_var+1

    # CONTINUITY EQUATION FOR EACH NODE
    for x_node in range(0, node.nbr_matrix.shape[0]):
        equ, equ_string = 0, ""
        for x_line in range(0, line.nbr_matrix.shape[0]):
            if var_sim.matrix_coupl_return[x_node, x_line] != 0:
                equ = equ+var_sim.matrix_coupl_return[x_node, x_line]*line.m_int_return[x_line]
                if var_sim.matrix_coupl_return[x_node, x_line] == 1:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" - "
                    else:
                        equ_string = "-"
                else:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" + "
                    else:
                        equ_string = "+"
                equ_string = equ_string+"m_int("+str(int(line.nbr_orig[x_line]))+")"
        if node.m_ext_return[x_node] != 0:
            equ = equ-node.m_ext_return[x_node]
            equ_string = equ_string+" - m_ext("+node.nbr_orig_roman[x_node]+")"
        else:
            equ_string = equ_string+" + 0"
        equ_string = equ_string+" = 0"
        equ_system_hydr_return_string.append(equ_string)
        equ_system_hydr_return.append(equ)

    # PRESSURE EQUATION FOR EACH LINE
    for x_line in range(0, line.nbr_matrix.shape[0]):
        equ, equ_string = 0, ""
        for x_node in range(0, node.nbr_matrix.shape[0]):
            if var_sim.matrix_coupl_return_trans[x_line, x_node] != 0:
                equ = equ+var_sim.matrix_coupl_return_trans[x_line, x_node]*(node.p_return[x_node]+var_H2O.rho*var_phy.g*node.h_coord[x_node])
                if var_sim.matrix_coupl_return_trans[x_line, x_node] == -1:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" - "
                    else:
                        equ_string = "-"
                else:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" + "
                equ_string = equ_string+"(p("+node.nbr_orig_roman[x_node]+") + rho * g * h("+node.nbr_orig_roman[x_node]+"))"
        equ = equ-((8*np.power(line.m_int_return[x_line], 2))/(np.power(line.dia[x_line], 4)*np.power(math.pi, 2)*var_H2O.rho))*(line.lambd[x_line]*(line.l[x_line]/line.dia[x_line])+line.zeta[x_line])
        equ_string = equ_string+" - 8 * m_int("+str(int(line.nbr_orig[x_line]))+")^2 / (dd("+str(int(line.nbr_orig[x_line]))+")^4 * pi^2 * rho) * (lambda("+str(int(line.nbr_orig[x_line]))+") * L("+str(int(line.nbr_orig[x_line]))+") / d("+\
            str(int(line.nbr_orig[x_line]))+") + zeta("+str(int(line.nbr_orig[x_line]))+"))"
        equ_string = equ_string + " = 0"
        equ_system_hydr_return_string.append(equ_string)
        equ_system_hydr_return.append(equ)
    
    return(equ_system_hydr_return)

###############################################################################
###############################################################################
# FUNCTION TO CREATE THE HYDRAULIC EQUATION SYSTEM FOR THE FORERUN ############
###############################################################################
###############################################################################
def equ_network_forerun(start_variables_forerun, line, node, var_sim, var_H2O, var_phy):
    """Creates the hydraulic equation system corresponding to the inputs 
    for the forerun: Continuity equations for each node and pressure equations 
    for each line.

    :param start_variables_forerun: Start variables for the equation system
    :type start_variables_forerun: numpy.ndarray

    :param line: Contains information about the lines
    :type line: line obj.

    :param node: Contains information about the nodes
    :type node: node obj.

    :param var_sim: Simulation variables
    :type var_sim: var_sim obj.

    :param var_H2O: Water variables
    :type var_H2O: var_H2O obj.

    :param var_phy: Physical variables
    :type var_phy: var_phy obj.

    :return: Hydraulic equation system
    :rtype: numpy.ndarray
    """    
    
    # INITIALIZE EQUATION SYSTEMS
    equ_system_hydr_forerun, equ_system_hydr_forerun_string = [], []

    # START VALUES FOR UNKNOWNS
    for x_line in range(0, line.nbr_matrix.shape[0]):
        line.m_int_forerun[x_line] = start_variables_forerun[x_line]
    cntr_var = x_line+1
    for x_node in range(0, node.nbr_matrix.shape[0]):
        if node.p_ref_forerun[x_node] == None:
            node.p_forerun[x_node] = start_variables_forerun[cntr_var]
            cntr_var = cntr_var+1
    for x_node in range(0, node.m_ext_forerun.shape[0]):
        if node.m_ext_forerun_check[x_node] == None:
            try:
                node.m_ext_forerun[x_node] = start_variables_forerun[cntr_var]
            except:
                print("Fehler")
            cntr_var = cntr_var+1

    # CONTINUITY EQUATION FOR EACH NODE
    for x_node in range(0, node.nbr_matrix.shape[0]):
        equ, equ_string = 0, ""
        for x_line in range(0, line.nbr_matrix.shape[0]):
            if var_sim.matrix_coupl_forerun[x_node, x_line] != 0:
                equ = equ+var_sim.matrix_coupl_forerun[x_node, x_line]*line.m_int_forerun[x_line]
                if var_sim.matrix_coupl_forerun[x_node, x_line] == 1:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" + "
                    else:
                        equ_string = "+"
                else:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" - "
                    else:
                        equ_string = "-"
                equ_string = equ_string+"m_int("+str(int(line.nbr_orig[x_line]))+")"
        if node.m_ext_forerun[x_node] != 0:
            equ = equ-node.m_ext_forerun[x_node]
            equ_string = equ_string+" - m_ext("+node.nbr_orig_roman[x_node]+")"
        else:
            equ_string = equ_string+" + 0"
        equ_system_hydr_forerun_string.append(equ_string)
        equ_system_hydr_forerun.append(equ)

    # PRESSURE EQUATION FOR EACH LINE
    for x_line in range(0, line.nbr_matrix.shape[0]):
        equ, equ_string = 0, ""
        for x_node in range(0, node.nbr_matrix.shape[0]):
            if var_sim.matrix_coupl_forerun_trans[x_line, x_node] != 0:
                equ = equ+var_sim.matrix_coupl_forerun_trans[x_line, x_node]*(node.p_forerun[x_node]+var_H2O.rho*var_phy.g*node.h_coord[x_node])
                if var_sim.matrix_coupl_forerun_trans[x_line, x_node] == -1:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" - "
                    else:
                        equ_string = "-"
                else:
                    if len(equ_string) != 0:
                        equ_string = equ_string+" + "
                equ_string = equ_string+"(p("+node.nbr_orig_roman[x_node]+") + rho * g * h("+node.nbr_orig_roman[x_node]+"))"
        equ = equ-((8*np.power(line.m_int_forerun[x_line], 2))/(np.power(line.dia[x_line], 4)*np.power(math.pi, 2)*var_H2O.rho))*(line.lambd[x_line]*(line.l[x_line]/line.dia[x_line])+line.zeta[x_line])
        equ_string = equ_string+" - 8 * m_int("+str(int(line.nbr_orig[x_line]))+")^2 / (d("+str(int(line.nbr_orig[x_line]))+")^4 * pi^2 * rho) * (lambda("+str(int(line.nbr_orig[x_line]))+") * L("+str(int(line.nbr_orig[x_line]))+") / d("+\
            str(int(line.nbr_orig[x_line]))+") + zeta("+str(int(line.nbr_orig[x_line]))+"))"
        equ_system_hydr_forerun_string.append(equ_string)
        equ_system_hydr_forerun.append(equ)
        
    return(equ_system_hydr_forerun)



# PLOTTING FUNCTION

def setup_or_clear_subplots(plots):
    '''
    Sets up subplots for visualization or clears them for new plots.

    :param plots: Contains plots and corresponding information
    :type plots: plots obj.

    :return: Axes for the plots and color bars
    :rtype: matplotlib.axes.Axes, matplotlib.axes.Axes, matplotlib.axes.Axes, matplotlib.axes.Axes
    '''
    if hasattr(plots, 'fig3') and plots.fig3.axes:
        # The figure exists and has axes. Clear them for new plots.
        for ax in plots.fig3.axes: 
            ax.clear()
    else:
        # Create the figure and subplots for the first time
        plots.fig3 = plt.figure(2, figsize=(21,9))
        plt.show(block=False)

    # Change active figure to fig3
    plt.figure(2)

    # Axes for Plot 1
    ax1 = plt.axes([0.03, 0.1, 0.4, 0.8])  # [left, bottom, width, height]

    # Color Bar 1 (next to Plot 1)
    cbar_ax1 = plt.axes([0.44, 0.1, 0.01, 0.8])  # Adjusted for a narrow color bar

    # Axes for Plot 2 (leaving some space after Color Bar 1)
    ax2 = plt.axes([0.51, 0.1, 0.4, 0.8])  # Adjust starting point and width similar to Plot 1

    # Color Bar 2 (next to Plot 2)
    cbar_ax2 = plt.axes([0.92, 0.1, 0.01, 0.8])  # Positioned at the right edge, similar to Color Bar 1
    
    # Explicitly remove existing colorbars
    for cbar in plots.colorbars:
        cbar.remove()
    plots.colorbars.clear()  # Clear the list after removing colorbars
    
    return ax1, ax2, cbar_ax1, cbar_ax2

def draw_graph(ax, matrix_coupl, pressures, node, line, plots):
    '''
    Draws the graph of the network with the pressures as node colors.

    :param ax: Axes object for the plot
    :type ax: matplotlib.axes.Axes

    :param matrix_coupl: Coupling matrix for the network
    :type matrix_coupl: numpy.ndarray

    :param pressures: Pressures at the nodes
    :type pressures: numpy.ndarray

    :param node: Contains information about the nodes
    :type node: node obj.

    :param line: Contains information about the lines
    :type line: line obj.

    :param plots: Contains plots and corresponding information
    :type plots: plots obj.

    :return: ScalarMap for the colorbar
    :rtype: matplotlib.cm.ScalarMappable
    '''
    G = nx.DiGraph()
    
    # Add nodes
    for x_node in range(node.nbr_matrix.shape[0]):
        # Check if there is a mass flow at the current node and time step
        if node.m_ext_forerun[x_node] != 0 or node.m_ext_return[x_node] != 0:
            mass_flow = True
        else:
            mass_flow = False

        # Check if node is a feeder
        if node.feed_in[x_node] != None:
            is_feeder = True
        else:
            is_feeder = False

        # Check if node is a distributor
        is_distributor = False
        if node.distrib[x_node] != None:
            if node.distrib[x_node] == "x":
                is_distributor = True
        G.add_node(node.nbr_orig_roman[x_node], 
                   pos=(node.x_coord[x_node], node.y_coord[x_node]), 
                   active=mass_flow, 
                   is_feeder=is_feeder,
                   is_distributor=is_distributor)
        
    
    # Add edges based on the coupling matrix
    for x_line in range(line.nbr_matrix.shape[0]):
        column = matrix_coupl[:, x_line]
        node_start, node_end = None, None
        for x_node in range(node.nbr_matrix.shape[0]):
            if column[x_node] == 1:
                node_start = node.nbr_orig_roman[x_node]
            elif column[x_node] == -1:
                node_end = node.nbr_orig_roman[x_node]
        if node_start and node_end:
            G.add_edge(node_start, node_end, diameter=line.dia[x_line], line_num=int(line.nbr_orig[x_line]))
    
    # Color nodes
    node_colors = [pressures[x_node]/1000 for x_node in range(node.nbr_matrix.shape[0])]
    norm = matplotlib.colors.Normalize(vmin=min(node_colors), vmax=max(node_colors))
    cmap = matplotlib.cm.viridis
    scalarMap = matplotlib.cm.ScalarMappable(norm=norm, cmap=cmap)
    
    # Draw the components of the graph
    pos = nx.get_node_attributes(G, 'pos')
    widths = [G[u][v]['diameter'] * 40 for u, v in G.edges()]

    active = nx.get_node_attributes(G, 'active')
    node_alphas = [1 if active[node] else 0.2 for node in G.nodes()]

    is_feeder = nx.get_node_attributes(G, 'is_feeder')
    node_linewidths = [4 if is_feeder[node] else 0 for node in G.nodes()]

    is_distributor = nx.get_node_attributes(G, 'is_distributor')
    node_sizes = [200 if is_distributor[node] else 400 for node in G.nodes()]

    for i, size in enumerate(node_sizes):
        if size == 400:
            # Scale area with mass flow
            size_factor = abs(node.m_ext_forerun[i]) * 1500
            size = 300 + size_factor
            node_sizes[i] = size

    
    nx.draw_networkx_nodes(G, pos, ax=ax, node_color=[scalarMap.to_rgba(x) for x in node_colors], cmap=cmap, node_size=node_sizes, alpha=node_alphas, linewidths=node_linewidths, edgecolors='red')
    nx.draw_networkx_labels(G, pos, ax=ax, font_size=8, font_color="white", font_weight="bold")
    nx.draw_networkx_edges(G, pos, ax=ax, arrows=True, arrowstyle="->", arrowsize=20, edge_color='black', width=widths)

    edge_labels = nx.get_edge_attributes(G, 'line_num')
    nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, ax=ax, font_color='black')
    
    # Add colorbar (use the figure to place it correctly)
    scalarMap.set_array(node_colors)
    return scalarMap
