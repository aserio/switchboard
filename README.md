<!--- 
Copyright (c) 2020-2021 Adrian S. Lemoine

Distributed under the Boost Software License, Version 1.0. 
(See accompanying file LICENSE_1_0.txt or copy at 
http://www.boost.org/LICENSE_1_0.txt)
--->

# Switchboard
Connecting MS Project to GitHub and Beyond

Switchboard is designed to import data from GitHub 
into a Microsoft Project file. It is made of two 
major components. A Python script which fetches 
data from GitHub and a VBA file which merges this 
data into the Microsoft Project file.

A configuration file is needed to show the 
relationship between the Microsoft Project File, 
the GitHub repository, and the length of a sprint 
for a particular project.