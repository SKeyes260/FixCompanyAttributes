# FixCompanyAttributes
Fixes SCCM client records whose AD value for companyAttributeMachineCategory does not match SCCM. Searches all company domains for matching computer names and creates an SCCM DDR file containing the computer Name, Distinguishedname and MachineCategory. If the SMSGUID and MachineType exist they are also included in the DDR
