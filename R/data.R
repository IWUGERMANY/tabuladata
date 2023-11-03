#' Info.Sheet
#' 
#' Internal sheet control
'info.sheet'

#' Const.DataFormat
#' 
#' Data format codes
'const.dataformat'

#' Const.StatusDataset
#' 
#' Dataset status codes
'const.statusdataset'

#' Const.TypeVariant
#' 
#' Type of building variant (refurbishment variant or new build variant)
'const.typevariant'

#' Const.TypeIntake
#' 
#' Types of data intake
'const.typeintake'

#' Const.DataType.Building
#' 
#' Building data type
'const.datatype.building'

#' Const.Country
#' 
#' Country codes
'const.country'

#' Const.Language
#' 
#' Language codes
'const.language'

#' Const.BuildingSizeClass
#' 
#' Building size classes
'const.buildingsizeclass'

#' Const.BoundaryCondType
#' 
#' Type of boundary condition input
'const.boundarycondtype'

#' Const.Utilisation
#' 
#' Common definitions of utilisation types
'const.utilisation'

#' Const.RoofType
#' 
#' Roof type codes
'const.rooftype'

#' Const.AtticCond
#' 
#' Space heating situation of the attic storey
'const.atticcond'

#' Const.CellarCond
#' 
#' Space heating situation of the cellar storey
'const.cellarcond'

#' Const.AttNeighb
#' 
#' Codes for the number of attached neighbour buildings
'const.attneighb'

#' Const.ComplexFootprint
#' 
#' Degree of geometrical complexity of the footprint shape
'const.complexfootprint'

#' Const.ComplexRoof
#' 
#' Degree of geometrical complexity of the roof shape
'const.complexroof'

#' Const.Orientation
#' 
#' Geographical orientation
'const.orientation'

#' Const.ThermalBridging
#' 
#' Classification of thermal bridging
'const.thermalbridging'

#' Const.Infiltration
#' 
#' Classification of infiltration, dependent on air tightness
'const.infiltration'

#' Const.ConstrBorder
#' 
#' Type of construction border situation / location of construction element
'const.constrborder'

#' Const.ElementType
#' 
#' Thermal envelope element types
'const.elementtype'

#' Const.MeasureType
#' 
#' Type of refurbishment measure / replacement of existing insulation or elements
'const.measuretype'

#' Const.EnergyCarrier
#' 
#' Codes for the used energy carriers
'const.energycarrier'

#' Const.SysType.HG
#' 
#' Heating system: generator types
'const.systype.hg'

#' Const.SysType.HS
#' 
#' Heating system: storage types
'const.systype.hs'

#' Const.SysType.HD
#' 
#' Heating system: distribution and heat emission types
'const.systype.hd'

#' Const.SysType.HA
#' 
#' Heating system: auxiliary energy types
'const.systype.ha'

#' Const.SysType.Vent
#' 
#' Heating system: ventilation types
'const.systype.vent'

#' Const.SysType.WG
#' 
#' Domestic hot water system: generator types
'const.systype.wg'

#' Const.SysType.WS
#' 
#' Domestic hot water system: storage types
'const.systype.ws'

#' Const.SysType.WD
#' 
#' Domestic hot water system: distribution types
'const.systype.wd'

#' Const.SysType.WA
#' 
#' Domestic hot water system: auxiliary energy types
'const.systype.wa'

#' Const.SysType.Size
#' 
#' Building types used to distinguish between different system component sizes
'const.systype.size'

#' Const.Type.CalcAdapt
#' 
#' Calculation adaptation types
'const.type.calcadapt'

#' TypologyRegion
#' 
#' National typology regions
'typologyregion'

#' ConstrYearClass
#' 
#' National construction year classes
'constryearclass'

#' AdditionalPar
#' 
#' National addtitional parameter for classification (regions, special building types etc.)
'additionalpar'

#' Uncertainty.Levels
#' 
#' Uncertainty levels: 5 categories of input data uncertainty; one row for each relevant input variable; for each relevant input variable: explanation of each category; quantification of absolute and/or relative uncertainty
'uncertainty.levels'

#' Climate
#' 
#' National and regional climate conditions
'climate'

#' BoundaryCond
#' 
#' Boundary conditions for the energy balance calculation (optional)
'boundarycond'

#' Par.EnvAreaEstim
#' 
#' Parameters for the envelope area estimation procedure
'par.envareaestim'

#' Building.Constr
#' 
#' National definition of construction elements + U-values
'building.constr'

#' Building.Measure
#' 
#' National definition of insulation measures + thermal resistance
'building.measure'

#' U.Class.Constr
#' 
#' National default U-values differentiated by construction year class and construction type
'u.class.constr'

#' Insulation.Default
#' 
#' National default values of insulation thickness and thermal conductivity of insulation measures, differentiated by implementation period and construction type; to be used if the actual values are not known for a given construction (usage of these default values increases the uncertainty)
'insulation.default'

#' Measure.f.Default
#' 
#' National default values of insulation fraction, differentiated by construction type, implementation period; two cases: (a) typical insulation fraction of refurbished construtions (indicator "Refurbished", used if actual insulation fraction is not known for a construction); (b) ratio of insulated to total construction area for the building stock of the specific construction period (indicator "NA", values used to model an average state for constructions of unknown state) - usage of these default values increases the uncertainty
'measure.f.default'

#' U.WindowType.Periods
#' 
#' National default U-values of windows, diferentiated by construction year class and type of glazing and frame
'u.windowtype.periods'

#' U.Class.Window
#' 
#' U-values of windows (old scheme used for TABULA WebTool (only three time bands)
'u.class.window'

#' System.HG
#' 
#' Heating system / heat generation default values
'system.hg'

#' System.HS
#' 
#' Heating system / heat storage default values
'system.hs'

#' System.HD
#' 
#' Heating system / heat distribution default values
'system.hd'

#' System.HA
#' 
#' Heating system / auxiliary energy default values
'system.ha'

#' System.WG
#' 
#' Domestic hot water system / heat generation default values
'system.wg'

#' System.WS
#' 
#' Domestic hot water system / heat storage default values
'system.ws'

#' System.WD
#' 
#' Domestic hot water system / heat distribution default values
'system.wd'

#' System.WA
#' 
#' Domestic hot water system / auxiliary energy default values
'system.wa'

#' System.H
#' 
#' National configurations of heating system
'system.h'

#' System.W
#' 
#' National configurations of domestic hot water systems
'system.w'

#' System.Vent
#' 
#' Default values of ventilation systems
'system.vent'

#' System.PVPanel
#' 
#' Default values for different PV panel types
'system.pvpanel'

#' System.PV
#' 
#' Default values of PV electricity production depending on orientation
'system.pv'

#' System.Coverage
#' 
#' Default values of coverages depending on supply/load ratio
'system.coverage'

#' System.ElProd
#' 
#' Assessment of on-site produced electricity
'system.elprod'

#' System.SetECAssess
#' 
#' Definition of energy carrier assessment sets
'system.setecassess'

#' System.EC
#' 
#' Energy carrier specification and assessment
'system.ec'

#' CalcAdapt
#' 
#' Definition of calibration factors for adapting the TABULA calculation results to the typical level of metered consumption or to the typical level of national calculation methods (used for respective estimations)
'calcadapt'
