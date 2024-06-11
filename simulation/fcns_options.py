"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

This script contains helper functions for the options file.

Functions:
- check_weather_file: Verifies the presence of a specified weather data file in a directory, 
  or returns the first available file if no specific file is requested.
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

import os

def check_weather_file(directory, file_name):
    '''
    Searches the specified path for csv files. If there is a file name specified,
    the function checks if the file is available in the directory. If not, the
    function returns the first file found in the directory.
    '''

    # List files in weather data directory
    weather_files = os.listdir(directory)
    # Filter for csv files
    weather_files = [file for file in weather_files if file.endswith('.csv')]
    # Sort list alphabetically
    weather_files.sort()

    # Check if a file name is specified
    if file_name:
        # Check if the file is available in the directory
        if not file_name in weather_files:
            raise Exception(f"Weather file {file_name} not found in directory {directory}.")
    else:
        # Return the first file found in the directory
        file_name = weather_files[0]

    file_path = os.path.join(directory, file_name)
    return file_path