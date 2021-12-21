<!--- 
Copyright (c) 2020-2021 Adrian S. Lemoine

Distributed under the Boost Software License, Version 1.0. 
(See accompanying file LICENSE_1_0.txt or copy at 
http://www.boost.org/LICENSE_1_0.txt)
--->

# Switchboard
Connecting MS Project to GitHub and Beyond

Switchboard is a project management toolkit 
designed to fetch data from external sources
such as GitHub and import it
into a Microsoft Project file. It is made of three 
major components: a VBA macro, a Python script, 
and a configuration file.

The VBA macro is responsible for the 
exectuion of the workflow. It reads in the
configuraion file, launches the
scripts necessary to fetch new data and merges this 
data into the Microsoft Project file.

The Python script, called a cord, 
fetches data from the external data source
and process it, and produces a CSV file 
for consumption by the VBA
macro.

The configuration file provides the VBA 
macro with the following information: 
 * The Python path
 * The path to the Switchboard installation
 * The name of the project
 * The source of the data
 * Other project infromation
   * Length of a sprint (Number of Days)
   * Pattern of the sprint name (default, "", or regular expression)