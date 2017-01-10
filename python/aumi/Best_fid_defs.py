#### SCHEDULE BA
PC_Unaffiliated =   ['SF21373', 'SF21375', 'SF21377', 'SF21379', 'SF21381', 'SF21383', 'SF21385', 'SF21387',
                     'SF21389', 'SF21391', 'SF21393', 'SF21395', 'SF21397', 'SF21399', 'SF21401', 'SF21403',
                     'SF21405', 'SF21407', 'SF21409', '', 'SF21411', 'SF21413', 'SF21414', '']
PC_Affiliated =     ['SF21374', 'SF21376', 'SF21378', 'SF21380', 'SF21382', 'SF21384', 'SF21386', 'SF21388',
                     'SF21390', 'SF21392', 'SF21394', 'SF21396', 'SF21398', 'SF21400', 'SF21402', 'SF21404',
                     'SF21406', 'SF21408', 'SF21410', '', 'SF21412', '', 'SF21415', '']
PC_Subcat =         [PC_Unaffiliated, PC_Affiliated]

LIFE_Unaffiliated = ['ST13332', 'ST13334', 'ST13336', 'ST13338', 'ST13340', 'ST13342', 'ST13344', 'ST13346',
                     'ST13348', 'ST25939', 'ST13350', 'ST13352', 'ST13354', 'ST16031', 'ST13356', 'ST15400',
                     'ST15402', 'ST25941', 'ST25943', '', 'ST15406', 'ST25945', 'ST13360', '']
LIFE_Affiliated =   ['ST13333', 'ST13335', 'ST13337', 'ST13339', 'ST13341', 'ST13343', 'ST13345', 'ST13347',
                     'ST13349', 'ST25940', 'ST13351', 'ST13353', 'ST13355', 'ST16032', 'ST13357', 'ST15401',
                     'ST15403', 'ST25942', 'ST25944', '', 'ST15407', '', 'ST13361', '']
LIFE_Subcat =       [LIFE_Unaffiliated, LIFE_Affiliated]

HEALTH_Unaffiliated =['HS08250', 'HS08252', 'HS08254', 'HS08256', 'HS08258', 'HS08260', 'HS08262', 'HS08264',
                      'HS08266', 'HS17035', 'HS08268', 'HS08270', 'HS08272', 'HS09299', 'HS08274', 'HS08809',
                      'HS08811', 'HS17037', 'HS17039', '', 'HS08815', 'HS17041', 'HS08278', '']
HEALTH_Affiliated =  ['HS08251', 'HS08253', 'HS08255', 'HS08257', 'HS08259', 'HS08261', 'HS08263', 'HS08265',
                      'HS08267', 'HS17036', 'HS08269', 'HS08271', 'HS08273', 'HS09300', 'HS08275', 'HS08810',
                      'HS08812', 'HS17038', 'HS17040', '', 'HS08816', '', 'HS08279', '']
HEALTH_Subcat =      [HEALTH_Unaffiliated, HEALTH_Affiliated]

#### FINANCIALS
PC_Iris_Ratios =    ['SF05352', 'SF05354', 'SF05356', 'SF05358', 'SF05360', 'SF05362', 'SF08471', 'SF08473',
                     'SF05366', 'SF05368', 'SF05370', 'SF05372', 'SF05374','','','','','','','','','','','','']
PC_Liquidity =      ['','', '','','', 'SF00025', 'SF00031', 'SF00043', 'SF00053', 'SF00139',
                     'SF07654', 'SF07656', 'SF00148', 'SF00147', 'SF00146']
PC_Net_Pos =        ['SF13840', 'SF00011', '', 'SF00013', 'SF00014', '', 'SF04762', 'SF04763', 'SF00016', 'SF06017',
                     'SF06018', 'SF04764', 'SF00023', 'SF00025','', 'SF00031', '', '', '']
PC_Fin_Sub_Tots =   [PC_Iris_Ratios, PC_Liquidity, PC_Net_Pos]

LIFE_Iris_Ratios =  ['','','','','','','','','','','','','','ST08607', 'ST08609', 'ST08611', 'ST08613',
                     'ST08615', 'ST08617', 'ST08619', 'ST08621', 'ST08623', 'ST08625', 'ST08627', 'ST08629']
LIFE_Liquidity =    ['','', '','','', 'ST00030', 'ST00031', 'ST00068', 'ST05266', 'ST00194',
                     'ST12180', 'ST12181', 'ST00213', 'ST00209', 'ST00208']
LIFE_Net_Pos =      ['ST19574', 'ST00016', '', 'ST00017', 'ST00018', '', 'ST07420', 'ST07421', 'ST00020', 'ST10039',
                     'ST10040', 'ST07422', 'ST00028', 'ST00030','ST12167', 'ST00031', 'ST00028', 'ST00043', 'ST00047']
LIFE_Fin_Sub_Tots =   [LIFE_Iris_Ratios, LIFE_Liquidity, LIFE_Net_Pos]

HEALTH_Iris_Ratios =  ['','','','','','','','','','','','','','', '', '', '',
                       '', '', '', '', '', '', '', '']
HEALTH_Liquidity =    ['','', '','','', 'HS00097', 'HS00100', '', '', 'HS00263',
                       'HS06177', 'HS06179', 'HS00271', 'HS00270', 'HS06180']
HEALTH_Net_Pos =      ['HS12695', 'HS00085', '', 'HS00086', 'HS00087', '', 'HS00088', 'HS00089','HS00090', 'HS00091', 'HS00092',
                       'HS00093', 'HS00094', 'HS00097','HS06143', 'HS00100', '', 'HS06151', 'HS00111']
HEALTH_Fin_Sub_Tots = [HEALTH_Iris_Ratios, HEALTH_Liquidity, HEALTH_Net_Pos]

### SCHEDULE D 1A 1
PC_US_Fed_Tot =     ['SF04565', 'SF04566', 'SF04567', 'SF04568', 'SF04569', 'SF04570']
PC_Oth_Gov_Tot =    ['SF10297', 'SF10298', 'SF10299', 'SF10300', 'SF10301', 'SF10302']
PC_US_St_Ter_Tot =  ['SF10304', 'SF10305', 'SF10306', 'SF10307', 'SF10308', 'SF10309']
PC_US_Pol_Sub_Tot = ['SF10311', 'SF10312', 'SF10313', 'SF10314', 'SF10315', 'SF10316']
PC_US_Spec_Rev_Tot =['SF10318', 'SF10319', 'SF10320', 'SF10321', 'SF10322', 'SF10323']
PC_Corp_Tot =       ['SF10325', 'SF10326', 'SF10327', 'SF10328', 'SF10329', 'SF10330']
PC_Hybrid_Tot =     ['SF10332', 'SF10333', 'SF10334', 'SF10335', 'SF10336', 'SF10337']
PC_Par_Sub_Aff_Tot =['SF02189', 'SF02190', 'SF02191', 'SF02192', 'SF02193', 'SF02194']
PC_Gov_Ag_Tot =     [PC_US_Fed_Tot, PC_Oth_Gov_Tot]
PC_Muni_Tot =       [PC_US_St_Ter_Tot, PC_US_Pol_Sub_Tot, PC_US_Spec_Rev_Tot]

PC_Combo_Asset_Classes = [PC_Gov_Ag_Tot, PC_Muni_Tot]
PC_Asset_Class_Tots =    [PC_US_Fed_Tot, PC_Oth_Gov_Tot, PC_US_St_Ter_Tot, PC_US_Pol_Sub_Tot,
                          PC_Gov_Ag_Tot, PC_Corp_Tot, PC_US_Spec_Rev_Tot, PC_Hybrid_Tot,
                          PC_Par_Sub_Aff_Tot, PC_Muni_Tot]

LIFE_US_Fed_Tot =        ['ST07163', 'ST07164', 'ST07165', 'ST07166', 'ST07167','ST07168']
LIFE_Oth_Gov_Tot =       ['ST18057', 'ST18058', 'ST18059', 'ST18060', 'ST18061', 'ST18062']
LIFE_US_St_Ter_Tot =     ['ST18064', 'ST18065', 'ST18066', 'ST18067', 'ST18068', 'ST18069']
LIFE_US_Pol_Sub_Tot =    ['ST18071', 'ST18072', 'ST18073', 'ST18074', 'ST18075', 'ST18076']
LIFE_US_Spec_Rev_Tot =   ['ST18078', 'ST18079', 'ST18080', 'ST18081', 'ST18082', 'ST18083']
LIFE_Corp_Tot =          ['ST18085', 'ST18086', 'ST18087', 'ST18088', 'ST18089', 'ST18090']
LIFE_Hybrid_Tot =        ['ST18092', 'ST18093', 'ST18094', 'ST18095', 'ST18096', 'ST18097']
LIFE_Par_Sub_Aff_Tot =   ['ST04422', 'ST04423', 'ST04424', 'ST04425', 'ST04426', 'ST04427']
LIFE_Gov_Ag_Tot =        [LIFE_US_Fed_Tot, LIFE_Oth_Gov_Tot]
LIFE_Muni_Tot =          [LIFE_US_St_Ter_Tot, LIFE_US_Pol_Sub_Tot, LIFE_US_Spec_Rev_Tot]

LIFE_Combo_Asset_Classes = [LIFE_Gov_Ag_Tot, LIFE_Muni_Tot]
LIFE_Asset_Class_Tots =    [LIFE_US_Fed_Tot, LIFE_Oth_Gov_Tot, LIFE_US_St_Ter_Tot, LIFE_US_Pol_Sub_Tot,
                            LIFE_Gov_Ag_Tot, LIFE_Corp_Tot, LIFE_US_Spec_Rev_Tot, LIFE_Hybrid_Tot,
                            LIFE_Par_Sub_Aff_Tot, LIFE_Muni_Tot]

HEALTH_US_Fed_Tot =        ['HS02710', 'HS02711', 'HS02712', 'HS02713', 'HS02714', 'HS02715']
HEALTH_Oth_Gov_Tot =       ['HS10291', 'HS10292', 'HS10293', 'HS10294', 'HS10295', 'HS10296']
HEALTH_US_St_Ter_Tot =     ['HS10298', 'HS10299', 'HS10300', 'HS10301', 'HS10302', 'HS10303']
HEALTH_US_Pol_Sub_Tot =    ['HS10305', 'HS10306', 'HS10307', 'HS10308', 'HS10309', 'HS10310']
HEALTH_US_Spec_Rev_Tot =   ['HS10312', 'HS10313', 'HS10314', 'HS10315', 'HS10316', 'HS10317']
HEALTH_Corp_Tot =          ['HS10319', 'HS10320', 'HS10321', 'HS10322', 'HS10323', 'HS10324']
HEALTH_Hybrid_Tot =        ['HS10326', 'HS10327', 'HS10328', 'HS10329', 'HS10330', 'HS10331']
HEALTH_Par_Sub_Aff_Tot =   ['HS03011', 'HS03012', 'HS03013', 'HS03014', 'HS03015', 'HS03016']
HEALTH_Gov_Ag_Tot =        [HEALTH_US_Fed_Tot, HEALTH_Oth_Gov_Tot]
HEALTH_Muni_Tot =          [HEALTH_US_St_Ter_Tot, HEALTH_US_Pol_Sub_Tot, HEALTH_US_Spec_Rev_Tot]

HEALTH_Combo_Asset_Classes = [HEALTH_Gov_Ag_Tot, HEALTH_Muni_Tot]
HEALTH_Asset_Class_Tots =    [HEALTH_US_Fed_Tot, HEALTH_Oth_Gov_Tot, HEALTH_US_St_Ter_Tot, HEALTH_US_Pol_Sub_Tot,
                              HEALTH_Gov_Ag_Tot, HEALTH_Corp_Tot, HEALTH_US_Spec_Rev_Tot, HEALTH_Hybrid_Tot,
                              HEALTH_Par_Sub_Aff_Tot, HEALTH_Muni_Tot]

LINES_PER_ASSET_CLASS = 10

ALL_Single_Data =['CO00231', '', 'CO00003', 'CO00179', '', 'CO00014', 'CO00030', 'CO00031',
                   'CO00033', 'CO00034', 'CO00035', '', 'CO00036', 'CO00020', 'CO00010',
                   'CO00327', 'CO00341', 'CO00324', 'CO00117', 'CO00022']
