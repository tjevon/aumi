# tag - used to locate filenames and template worksheets
SI01_tag     = "SI01"
E07_tag      = "E07"
BA_Acq_tag   = "BA_Acq"
BA_Disp_tag  = "BA_Disp"
E10_tag      = "E10"
Liab1_tag    = "Liab1"
Liab2_tag    = "Liab2"
Liab3_tag    = "Liab3"
SI05_07_tag  = "SI05_07"
SI08_09_tag  = "SI08_09"
Assets_tag   = "Assets"
SoI_tag      = "SoI"
SoO_tag      = "SoO"
SoR_tag      = "SoR"
IRIS1_tag    = "IRIS1"
IRIS2_tag    = "IRIS2"
CashFlow_tag = "CashFlow"
CR_tag       = "CR"
Liquid_Assets_tag = "Liquid_Assets"
Liq_Acq_tag = "Liq_Acq"
Liq_Disp_tag = "Liq_Disp"
Asset_Alloc_tag = "Asset_Alloc"
Real_Estate_tag = "Real_Estate"

PC_tag      = "PC"
LIFE_tag    = "LIFE"
HEALTH_tag  = "HEALTH"

COMMON_TEMPLATE_TAGS = []
COMMON_QUARTERLY_TAGS = []

ANNUAL = False

if ANNUAL:
    COMMON_TEMPLATE_TAGS = [Assets_tag, Liquid_Assets_tag, E07_tag, Real_Estate_tag,
                            CashFlow_tag, SI05_07_tag, E10_tag, Asset_Alloc_tag]
    COMMON_QUARTERLY_TAGS = [Assets_tag, Liquid_Assets_tag,  E07_tag, CashFlow_tag,
                             Real_Estate_tag]
else:
    COMMON_TEMPLATE_TAGS = [Assets_tag, Liquid_Assets_tag, Liq_Acq_tag, Liq_Disp_tag,
                            BA_Acq_tag, BA_Disp_tag, E07_tag, Real_Estate_tag,
                            CashFlow_tag, SI05_07_tag, E10_tag, Asset_Alloc_tag]
    COMMON_QUARTERLY_TAGS = [Assets_tag, Liquid_Assets_tag, Liq_Acq_tag, Liq_Disp_tag,
                             BA_Acq_tag, BA_Disp_tag, E07_tag, CashFlow_tag,
                             Real_Estate_tag]

# Needed for Quarterly report
PC_TEMPLATE_TAGS = [SoI_tag]
LIFE_TEMPLATE_TAGS = [SoO_tag]
HEALTH_TEMPLATE_TAGS = [SoR_tag]

PC_QUARTERLY_TAGS = [SoI_tag]
LIFE_QUARTERLY_TAGS = [SoO_tag]
HEALTH_QUARTERLY_TAGS = [SoR_tag]


