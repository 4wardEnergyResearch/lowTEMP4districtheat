# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

A wrapper for the data preparation subroutines.
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

import termcolor
import options as options
import importlib

# RELOAD OPTIONS ##############################################################
###############################################################################
# Reload options to make sure that the latest changes are applied.
importlib.reload(options)

def print_magenta(text):
    print(termcolor.colored(text, 'magenta'))

print_magenta("Executing data prep.")
print_magenta("")
print_magenta("Extracting consumer info...")
# NOTE: This routine may have to be adapted to your specific data structure.
from data_prep import data_prep_consumer_list_analysis
print_magenta("Preparing consumer data...")
# NOTE: This routine may have to be adapted to your specific data structure.
from data_prep import data_prep_consumers
print_magenta("Preparing feeder data...")
# NOTE: This routine may have to be adapted to your specific data structure.
from data_prep import data_prep_feeders
print_magenta("Adjusting SLPs...")
from data_prep import data_prep_load_profiles
print_magenta("Done.")

