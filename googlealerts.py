import os
import re
import xml
import xlwt
import requests
import html2text
from datetime import datetime
from bs4 import BeautifulSoup
from EmailDigestAPI import EmailDigestAPI

LA_HONDA_ALERTS_URLS = {

#VAR Target List New Beg	
#"https://www.google.com/alerts/feeds/13466654299695330520/477734780287099036": "Fusionstorm", 
#"https://www.google.com/alerts/feeds/13466654299695330520/2889460546644572078": "Trace3", 
#"https://www.google.com/alerts/feeds/13466654299695330520/3626046738560358261": "Denali Advanced Integration", 
#"https://www.google.com/alerts/feeds/13466654299695330520/18299370789804107852": "Integrated Archive Systems", 
#"https://www.google.com/alerts/feeds/13466654299695330520/4741467760900212027": "GroupWare", 
#"https://www.google.com/alerts/feeds/13466654299695330520/6029716097834067331": "Blue Chip Tek", 
#"https://www.google.com/alerts/feeds/13466654299695330520/6029716097834068590": "Bedrock Technology Partners", 
#"https://www.google.com/alerts/feeds/13466654299695330520/8665283408024576059": "Enterprise Vision Technologies",
#"https://www.google.com/alerts/feeds/13466654299695330520/11384353388350349739": "Entisys360", 
#"https://www.google.com/alerts/feeds/13466654299695330520/12743713329663882141": "Kovarus", 
#"https://www.google.com/alerts/feeds/13466654299695330520/12743713329663882106": "Insight Investments", 
#"https://www.google.com/alerts/feeds/13466654299695330520/6377559386693398851": "Dasher Technologies", 
#"https://www.google.com/alerts/feeds/13466654299695330520/9973157840582836927": "ECS Imaging Inc.", 
#"https://www.google.com/alerts/feeds/13466654299695330520/9973157840582839109": "Nth Generation Computing", 
#"https://www.google.com/alerts/feeds/13466654299695330520/11276258610200309812": "Zones", 
#"https://www.google.com/alerts/feeds/13466654299695330520/11325695364949154161": "PC Connection", 
#"https://www.google.com/alerts/feeds/13466654299695330520/13987126433168920890": "Softchoice", 
#"https://www.google.com/alerts/feeds/13466654299695330520/5810579531890136314": "World Wide Technology", 
#"https://www.google.com/alerts/feeds/13466654299695330520/13557375106724817821": "3RP Solutions", 
#"https://www.google.com/alerts/feeds/13466654299695330520/11962719138877261432": "Business IT Source", 
#"https://www.google.com/alerts/feeds/13466654299695330520/5492712225867939228": "eGroup",
#"https://www.google.com/alerts/feeds/13466654299695330520/18304040471736823950": "High Point Networks", 
#"https://www.google.com/alerts/feeds/13466654299695330520/18304040471736824118": "Myriad Supply", 
#"https://www.google.com/alerts/feeds/13466654299695330520/513645705325212977": "Carahsoft", 
#"https://www.google.com/alerts/feeds/13466654299695330520/9631845521440810943": "Avanade",
#"https://www.google.com/alerts/feeds/13466654299695330520/4760693138674747819": "Iron Bow Technologies",
#"https://www.google.com/alerts/feeds/13466654299695330520/8050474147697364694": "CcIntegration", 
#"https://www.google.com/alerts/feeds/13466654299695330520/860697833214560117": "Layer 3 Communications", 
#"https://www.google.com/alerts/feeds/13466654299695330520/10729875605318364586": "Hawk Ridge Systems", 
#"https://www.google.com/alerts/feeds/13466654299695330520/6581263688812950116": "Presidio", 
#VAR Target List New End

#NEW Beg
"https://www.google.com/alerts/feeds/13466654299695330520/10051877883879485662": "3DCart", 
: "3xLOGIC", 
: "Access Softek", 
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",
: "",


#NEW end










"https://www.google.com/alerts/feeds/02062406705762041312/5451535100808539110": "Zinio",
"https://www.google.com/alerts/feeds/02062406705762041312/17347623085535544058": "Automic",
"https://www.google.com/alerts/feeds/02062406705762041312/15671211174766790806": "Bizagi",
"https://www.google.com/alerts/feeds/020624067057620413121/5451535100808537977": "User Testing",
"https://www.google.com/alerts/feeds/02062406705762041312/13734466308953168108":    "Softfacade",
"https://www.google.com/alerts/feeds/02062406705762041312/13734466308953168829":    "Shotspotter",
"https://www.google.com/alerts/feeds/02062406705762041312/11092147724755190784":    "scandigital",
"https://www.google.com/alerts/feeds/02062406705762041312/2656591414836911737": "Red Pine Signals",
"https://www.google.com/alerts/feeds/02062406705762041312/11893550640734866279":    "Red Giant",
"https://www.google.com/alerts/feeds/02062406705762041312/11893550640734866364":    "Pley",
# "https://www.google.com/alerts/feeds/02062406705762041312/4965430515445542089": "Plex",
"https://www.google.com/alerts/feeds/02062406705762041312/4965430515445539778": "PaperG",
"https://www.google.com/alerts/feeds/02062406705762041312/11893550640734865632":    "Mobile Defense",
"https://www.google.com/alerts/feeds/02062406705762041312/18127349940471102005":    "JadooTV",
"https://www.google.com/alerts/feeds/02062406705762041312/1842992242340862655": "Fuzz Productions",
"https://www.google.com/alerts/feeds/02062406705762041312/5769740845327470898": "Edifecs",
"https://www.google.com/alerts/feeds/02062406705762041312/5769740845327470649": "APX, Inc.",
"https://www.google.com/alerts/feeds/02062406705762041312/14443331568096937893":    "Zoom Caffe",
"https://www.google.com/alerts/feeds/02062406705762041312/14852976520710281506":    "Yardi",
"https://www.google.com/alerts/feeds/02062406705762041312/14443331568096940018":    "WRS Materials",
"https://www.google.com/alerts/feeds/02062406705762041312/8615484252419376638": "Well.ca",
"https://www.google.com/alerts/feeds/02062406705762041312/8615484252419377547": "Votigo",
"https://www.google.com/alerts/feeds/02062406705762041312/8615484252419378175": "vitria",
"https://www.google.com/alerts/feeds/02062406705762041312/8933258069582921537": "VisiQuate",
"https://www.google.com/alerts/feeds/02062406705762041312/5746718541562635731": "Virtuos Ltd.",
"https://www.google.com/alerts/feeds/02062406705762041312/5746718541562635919": "Virool",
"https://www.google.com/alerts/feeds/02062406705762041312/18334767804339456745":    "VigilNet Community Monitoring",
"https://www.google.com/alerts/feeds/02062406705762041312/5701420191287790796": "Vertical Systems",
"https://www.google.com/alerts/feeds/02062406705762041312/4021406299706766443": "Unbounce",
"https://www.google.com/alerts/feeds/02062406705762041312/8233567744689238324": "Ubiquity Global Services",
"https://www.google.com/alerts/feeds/02062406705762041312/4021406299706767568": "Tutela Tech",
"https://www.google.com/alerts/feeds/02062406705762041312/12122000424484240313":    "Tulip Retail",
"https://www.google.com/alerts/feeds/02062406705762041312/12122000424484240178":    "Thoughtworks",
"https://www.google.com/alerts/feeds/02062406705762041312/12122000424484239265":    "Optime Group",
"https://www.google.com/alerts/feeds/02062406705762041312/5740363056375727398": "Tasktop",
"https://www.google.com/alerts/feeds/02062406705762041312/5740363056375728928": "Talend",
"https://www.google.com/alerts/feeds/02062406705762041312/3722266231542023458": "Systems in Motion",
"https://www.google.com/alerts/feeds/02062406705762041312/3722266231542023054": "Survey Analytics",
"https://www.google.com/alerts/feeds/02062406705762041312/10682813675916038701":    "Sureline Sytems",
"https://www.google.com/alerts/feeds/02062406705762041312/1782326813927721914": "SugarCRM",
"https://www.google.com/alerts/feeds/02062406705762041312/632202047199448876":  "stat Health Services",
"https://www.google.com/alerts/feeds/02062406705762041312/3088495681118562133": "Smule",
"https://www.google.com/alerts/feeds/02062406705762041312/14607657986943534996":    "SMS, inc.",
"https://www.google.com/alerts/feeds/02062406705762041312/7903596616217591861": "SkyTree",
"https://www.google.com/alerts/feeds/02062406705762041312/7903596616217592967": "Sitecore",
"https://www.google.com/alerts/feeds/02062406705762041312/12883977902754348340":    "Signiant",
"https://www.google.com/alerts/feeds/02062406705762041312/12883977902754345668":    "Sift Shopping",
"https://www.google.com/alerts/feeds/02062406705762041312/10010665971850991801":    "Schedulicity",
"https://www.google.com/alerts/feeds/02062406705762041312/10010665971850993571":    "Rolith",
"https://www.google.com/alerts/feeds/02062406705762041312/9919732738508487404": "Roamware",
"https://www.google.com/alerts/feeds/02062406705762041312/692729505549121703":  "Referral Saasquatch",
"https://www.google.com/alerts/feeds/02062406705762041312/692729505549121730":  "RedBubble",
"https://www.google.com/alerts/feeds/02062406705762041312/17193276624174069073":    "Recorded Future",
"https://www.google.com/alerts/feeds/02062406705762041312/17193276624174068920":    "Recon Instruments",
"https://www.google.com/alerts/feeds/02062406705762041312/8436229574644343393": "Recommind",
"https://www.google.com/alerts/feeds/02062406705762041312/1759710790041120949": "Realty Mogul",
"https://www.google.com/alerts/feeds/02062406705762041312/6991931469844997335": "Radiant Logic",
"https://www.google.com/alerts/feeds/02062406705762041312/6991931469844996642": "QuickMobile",
"https://www.google.com/alerts/feeds/02062406705762041312/11257665490522839832":    "Quantisense",
"https://www.google.com/alerts/feeds/02062406705762041312/11219233900047042895":    "Pulson",
"https://www.google.com/alerts/feeds/02062406705762041312/10429503542873433988":    "Promevo",
"https://www.google.com/alerts/feeds/02062406705762041312/10429503542873433334":    "Proformative",
"https://www.google.com/alerts/feeds/02062406705762041312/5240033680261613147": "prevoty",
"https://www.google.com/alerts/feeds/02062406705762041312/5771414264762800165": "Pretio Interactive",
"https://www.google.com/alerts/feeds/02062406705762041312/14135372739064166543":    "Plum Voice",
"https://www.google.com/alerts/feeds/02062406705762041312/164667081325941276":  "Percona",
"https://www.google.com/alerts/feeds/02062406705762041312/9070022432653348621": "PDHI",
"https://www.google.com/alerts/feeds/02062406705762041312/9070022432653348074": "Oversight Systems",
"https://www.google.com/alerts/feeds/02062406705762041312/17655246754798696192":    "Ontraport",
"https://www.google.com/alerts/feeds/02062406705762041312/6973255723239729332": "Neusoft",
"https://www.google.com/alerts/feeds/02062406705762041312/107409442514440098":  "Netwrix",
"https://www.google.com/alerts/feeds/02062406705762041312/11610665288912776873":    "neato",
"https://www.google.com/alerts/feeds/02062406705762041312/3508825709076618444": "navagate",
"https://www.google.com/alerts/feeds/02062406705762041312/3508825709076616710": "mycorporation.com",
"https://www.google.com/alerts/feeds/02062406705762041312/3508825709076616997": "Mulesoft",
"https://www.google.com/alerts/feeds/02062406705762041312/12931894682642024650":    "Mobile Action",
"https://www.google.com/alerts/feeds/02062406705762041312/6981264731603356147": "Malwarebytes",
"https://www.google.com/alerts/feeds/02062406705762041312/15191979296832764568":    "Main Street Hub",
"https://www.google.com/alerts/feeds/02062406705762041312/10156193641003657413":    "Lumo Bodytech",
"https://www.google.com/alerts/feeds/02062406705762041312/15613636266159631161":    "Odyssey Entertainment",
"https://www.google.com/alerts/feeds/02062406705762041312/17404663272574942251":    "LiveEnsure",
"https://www.google.com/alerts/feeds/02062406705762041312/1390794130710317575": "Lithium Technologies",
"https://www.google.com/alerts/feeds/02062406705762041312/17391609790282109399":    "Line2",
"https://www.google.com/alerts/feeds/02062406705762041312/17396427411805734271":    "Kubicam",
"https://www.google.com/alerts/feeds/02062406705762041312/6575616247985574918": "Knowledge Marketing",
"https://www.google.com/alerts/feeds/02062406705762041312/17396427411805735603":    "KeyedIn Solutions",
"https://www.google.com/alerts/feeds/02062406705762041312/3577793260310607629": "joyent",
"https://www.google.com/alerts/feeds/02062406705762041312/200946860401571985":  "jamcracker",
"https://www.google.com/alerts/feeds/02062406705762041312/200946860401570976":  "interneer",
"https://www.google.com/alerts/feeds/02062406705762041312/8608874187962445793": "internet roi",
"https://www.google.com/alerts/feeds/02062406705762041312/3826162343940337764": "integrated biometrics",
"https://www.google.com/alerts/feeds/02062406705762041312/6223083195894846682": "ins zoom",
"https://www.google.com/alerts/feeds/02062406705762041312/16776894766737984906":    "inriver",
"https://www.google.com/alerts/feeds/02062406705762041312/16776894766737983124":    "information builders",
"https://www.google.com/alerts/feeds/02062406705762041312/11127422907426619373":    "icims",
"https://www.google.com/alerts/feeds/02062406705762041312/2962025794278729148": "gyrus",
"https://www.google.com/alerts/feeds/02062406705762041312/11127422907426618749":    "gemini solutions",
"https://www.google.com/alerts/feeds/02062406705762041312/7395578459586164283": "fullarmor",
"https://www.google.com/alerts/feeds/02062406705762041312/13247670572481404415":    "freshbooks",
"https://www.google.com/alerts/feeds/02062406705762041312/109613453342848590":  "fracturedme.com",
"https://www.google.com/alerts/feeds/02062406705762041312/109613453342847149":  "four winds interactive",
"https://www.google.com/alerts/feeds/02062406705762041312/1914311863843569760": "flowgear",
"https://www.google.com/alerts/feeds/02062406705762041312/16279816956832154522":    "etwater",
"https://www.google.com/alerts/feeds/02062406705762041312/16279816956832155729":    "ericom",
"https://www.google.com/alerts/feeds/02062406705762041312/16989223789975808488": "Clearink",
"https://www.google.com/alerts/feeds/02062406705762041312/14601896837141013602": "Contenix",
"https://www.google.com/alerts/feeds/02062406705762041312/1126258742458971627": "Moment Design",
"https://www.google.com/alerts/feeds/02062406705762041312/5657753612956061956": "SunCentral",
"https://www.google.com/alerts/feeds/02062406705762041312/9204548976826551664": "XG sciences",
"https://www.google.com/alerts/feeds/02062406705762041312/12510595579474040014": "Intervision",
"https://www.google.com/alerts/feeds/02062406705762041312/2072179285307919018": "Kiip",
"https://www.google.com/alerts/feeds/02062406705762041312/7019550178776333209": "Sojern",
"https://www.google.com/alerts/feeds/02062406705762041312/7019550178776333510": "Xirrus",
"https://www.google.com/alerts/feeds/02062406705762041312/12011210292130282520": "Allocadia",
"https://www.google.com/alerts/feeds/02062406705762041312/12011210292130281856": "Binwise",
"https://www.google.com/alerts/feeds/02062406705762041312/3195075836423366027": "Gener8",
"https://www.google.com/alerts/feeds/02062406705762041312/11660103569085035494": "Meridian Clean Coal",
"https://www.google.com/alerts/feeds/02062406705762041312/6392144421683957432": "Mocana",
"https://www.google.com/alerts/feeds/02062406705762041312/12011210292130282520": "Proxio",
"https://www.google.com/alerts/feeds/02062406705762041312/14443331568096937640": "Starview",
"https://www.google.com/alerts/feeds/02062406705762041312/8615484252419376870": "Stratogent",
"https://www.google.com/alerts/feeds/02062406705762041312/15632915224957884581": "TurnCommerce",
"https://www.google.com/alerts/feeds/02062406705762041312/5746718541562638237": "Vionx",
"https://www.google.com/alerts/feeds/02062406705762041312/15632915224957882477": "Virtual Bridges",
"https://www.google.com/alerts/feeds/02062406705762041312/18334767804339457246": "ViZn Energy",
"https://www.google.com/alerts/feeds/02062406705762041312/6661669010023826466": "Zoom Technologies",
"https://www.google.com/alerts/feeds/02062406705762041312/8233567744689240336": "ActMobile",
"https://www.google.com/alerts/feeds/02062406705762041312/4021406299706765511": "Armanta",
"https://www.google.com/alerts/feeds/02062406705762041312/6751348586596979261": "Leeyo",
"https://www.google.com/alerts/feeds/02062406705762041312/5238135823142396240": "Q1Media",
"https://www.google.com/alerts/feeds/02062406705762041312/5238135823142398697": "Search Technologies",
"https://www.google.com/alerts/feeds/02062406705762041312/5740363056375729869": "Secure64",
"https://www.google.com/alerts/feeds/02062406705762041312/16371578590400077312": "Soraa",
"https://www.google.com/alerts/feeds/02062406705762041312/3722266231542024170": "Spotify",
"https://www.google.com/alerts/feeds/02062406705762041312/10682813675916039401": "TigerText",
"https://www.google.com/alerts/feeds/02062406705762041312/17365579478526992457": "Elk River Systems",
"https://www.google.com/alerts/feeds/02062406705762041312/1782326813927722798": "Scribblelive",
"https://www.google.com/alerts/feeds/02062406705762041312/5228429197450360129": "360pi",
"https://www.google.com/alerts/feeds/02062406705762041312/3088495681118562550": "ActivEngage",
"https://www.google.com/alerts/feeds/02062406705762041312/2176721985226159447": "Act-On Software",
"https://www.google.com/alerts/feeds/02062406705762041312/2176721985226160603": "Adaptive Planning",
"https://www.google.com/alerts/feeds/02062406705762041312/7903596616217593985": "AppLovin",
"https://www.google.com/alerts/feeds/02062406705762041312/7833637987667539397": "Ayla Networks",
"https://www.google.com/alerts/feeds/02062406705762041312/7903596616217594296": "BigML",
"https://www.google.com/alerts/feeds/02062406705762041312/12766643275826260061": "BuildDirect",
"https://www.google.com/alerts/feeds/02062406705762041312/10488344716461284916": "CardCompliant",
"https://www.google.com/alerts/feeds/02062406705762041312/12766643275826258376": "Cavendish Kinetics",
"https://www.google.com/alerts/feeds/02062406705762041312/3811454850071931956": "Comodo",
# "https://www.google.com/alerts/feeds/02062406705762041312/617096218391768433": "Cooper",
"https://www.google.com/alerts/feeds/02062406705762041312/9919732738508485209": "Couchbase",
"https://www.google.com/alerts/feeds/02062406705762041312/4350246155083242847": "Coupa Software",
"https://www.google.com/alerts/feeds/02062406705762041312/9454263677642889586": "EngagePoint",
"https://www.google.com/alerts/feeds/02062406705762041312/17193276624174070004": "Enterprise Engineering",
"https://www.google.com/alerts/feeds/02062406705762041312/8436229574644345338": "eShipGlobal",
"https://www.google.com/alerts/feeds/02062406705762041312/8436229574644345728": "Fuhu",
# "https://www.google.com/alerts/feeds/02062406705762041312/2718969084101024173": "Good Technology",
"https://www.google.com/alerts/feeds/02062406705762041312/2718969084101022743": "HeartMath",
"https://www.google.com/alerts/feeds/02062406705762041312/12930299050015098788": "iboss",
"https://www.google.com/alerts/feeds/02062406705762041312/2537868182094187908": "Infolinks",
"https://www.google.com/alerts/feeds/02062406705762041312/10429503542873435873": "Inside Vault",
"https://www.google.com/alerts/feeds/02062406705762041312/9845147377187965561": "Kontiki",
"https://www.google.com/alerts/feeds/02062406705762041312/14596073740068426969": "Meltwater",
"https://www.google.com/alerts/feeds/02062406705762041312/12317356517016336115": "Mobify",
"https://www.google.com/alerts/feeds/02062406705762041312/3368120317282496202": "Outsystems",
"https://www.google.com/alerts/feeds/02062406705762041312/11041838645369549994": "PetersenDean",
"https://www.google.com/alerts/feeds/02062406705762041312/14534779934515597310": "Quantifind",
"https://www.google.com/alerts/feeds/02062406705762041312/11697718741169257071": "Quixey",
"https://www.google.com/alerts/feeds/02062406705762041312/9798962351390451465": "RGB Spectrum",
"https://www.google.com/alerts/feeds/02062406705762041312/14189370591100224223": "Rimini Street",
"https://www.google.com/alerts/feeds/02062406705762041312/11678059886279920551": "Rise Interactive",
"https://www.google.com/alerts/feeds/02062406705762041312/14189370591100224118": "Real-Time Innovations",
"https://www.google.com/alerts/feeds/02062406705762041312/15634272773826690271": "ShareThrough",
"https://www.google.com/alerts/feeds/02062406705762041312/1952998804414674328": "ShippingEasy",
"https://www.google.com/alerts/feeds/02062406705762041312/4248703195252150694": "TraceSecurity",
"https://www.google.com/alerts/feeds/02062406705762041312/17870585110378373068": "Trueffect",
# "https://www.google.com/alerts/feeds/02062406705762041312/9959723930708512576": "Turn",
"https://www.google.com/alerts/feeds/02062406705762041312/6040473229954067858": "Adconion",
"https://www.google.com/alerts/feeds/02062406705762041312/14199054692078207438": "VoiceBox Technologies",
"https://www.google.com/alerts/feeds/02062406705762041312/8541155846467231691": "Xamarin",
"https://www.google.com/alerts/feeds/02062406705762041312/16282961076157299363": "zSpace",
"https://www.google.com/alerts/feeds/02062406705762041312/8829297520172324704": "Appthority",
"https://www.google.com/alerts/feeds/02062406705762041312/16763980993319175496": "Sikka Software",
"https://www.google.com/alerts/feeds/02062406705762041312/9088983416957814343": "3esi",
"https://www.google.com/alerts/feeds/02062406705762041312/866446058370717187": "Acclivity",
"https://www.google.com/alerts/feeds/02062406705762041312/18391294624796254": "Acquia",
"https://www.google.com/alerts/feeds/02062406705762041312/16461820305707803567": "Adexa",
"https://www.google.com/alerts/feeds/02062406705762041312/3245231855590886604": "Aerospike",
"https://www.google.com/alerts/feeds/02062406705762041312/14847574306059095572": "Aircuity",
"https://www.google.com/alerts/feeds/02062406705762041312/13314189900806450833": "AlienVault",
"https://www.google.com/alerts/feeds/02062406705762041312/8043453079708019667": "Amperics",
"https://www.google.com/alerts/feeds/02062406705762041312/528568041899009378": "Anthem Media Group",
"https://www.google.com/alerts/feeds/02062406705762041312/5471179202261382482": "Apateq",
"https://www.google.com/alerts/feeds/02062406705762041312/18208327643445941080": "Appointment-Plus",
"https://www.google.com/alerts/feeds/02062406705762041312/18208327643445943061": "Artec Group",
"https://www.google.com/alerts/feeds/02062406705762041312/8197277581016720571": "Asigra",
"https://www.google.com/alerts/feeds/02062406705762041312/17036040936133735920": "ASSIA",
"https://www.google.com/alerts/feeds/02062406705762041312/8810453785864400720": "Atlantis Computing",
"https://www.google.com/alerts/feeds/02062406705762041312/3815617076212362936": "Automation Anywhere",
"https://www.google.com/alerts/feeds/02062406705762041312/12664269602168331769": "Azul Systems",
"https://www.google.com/alerts/feeds/02062406705762041312/8448563982377872067": "BenBria",
"https://www.google.com/alerts/feeds/02062406705762041312/12469002039915996178": "black lotus communications",
"https://www.google.com/alerts/feeds/02062406705762041312/4939292664339777742": "Bloomfire",
"https://www.google.com/alerts/feeds/02062406705762041312/12219022977072856547": "Blue Jeans Network",
"https://www.google.com/alerts/feeds/02062406705762041312/1081338962497421778": "Blue Wave Media",
"https://www.google.com/alerts/feeds/02062406705762041312/16585002916782872318": "BlueCat Networks",
"https://www.google.com/alerts/feeds/02062406705762041312/14088811267400607562": "Borealis",
"https://www.google.com/alerts/feeds/02062406705762041312/11303193489297562381": "Brainspace",
"https://www.google.com/alerts/feeds/02062406705762041312/13328510058997422486": "Chaordix",
"https://www.google.com/alerts/feeds/02062406705762041312/7397566731362060032": "ChapterThree",
"https://www.google.com/alerts/feeds/02062406705762041312/7638990815327381860": "Comilion",
"https://www.google.com/alerts/feeds/02062406705762041312/12487352408347965472": "Continuum Analytics",
"https://www.google.com/alerts/feeds/02062406705762041312/2600888525683738979": "Corel",
"https://www.google.com/alerts/feeds/02062406705762041312/15618437625746337895": "CTC America",
"https://www.google.com/alerts/feeds/02062406705762041312/57166365469168311": "David Corporation",
"https://www.google.com/alerts/feeds/02062406705762041312/16479446172795132828": "Digital Defense",
"https://www.google.com/alerts/feeds/02062406705762041312/9802622504040440941": "Digital Dream Labs",
"https://www.google.com/alerts/feeds/02062406705762041312/896319013666807366": "Double Line Partner's Education",
"https://www.google.com/alerts/feeds/02062406705762041312/2308958562630037518": "Dwell Media",
"https://www.google.com/alerts/feeds/02062406705762041312/9615978415870236504": "Echo Sec",
"https://www.google.com/alerts/feeds/02062406705762041312/2578744933996684994": "Elastic Path",
"https://www.google.com/alerts/feeds/02062406705762041312/9342015003890982679": "Encepta",
"https://www.google.com/alerts/feeds/02062406705762041312/8169993311602351141": "Engine Yard",
"https://www.google.com/alerts/feeds/02062406705762041312/2641415568862031826": "Enprecis",
"https://www.google.com/alerts/feeds/02062406705762041312/6503229101590592373": "Electric Cloud"
}

def removeHTMLTags(data):
	  p = re.compile(r'<.*?>')
	  return p.sub('', data)

def cleanTitle(title):
	  title = removeHTMLTags(str(title))
	  title = title.replace('&lt;b&gt;', '')
	  title = title.replace('&lt;/b&gt;', '')
	  title = title.replace('&amp;#39;', '\'')
	  return title

def cleanLink(link):
	  link = str(link)
	  start = link.find('url=')
	  link = link[start + 4:]
	  end = link.find('&amp')
	  link = link[:end]
	  return link

def cleanDate(date):
	  date = removeHTMLTags(str(date))
	  end = date.find('T')
	  date = date[:end]
	  return date

def writeToSheet(sheet, title, company, link, date, cur_row):
	  sheet.write(cur_row, 3, unicode(title, "utf-8"))
	  sheet.write(cur_row, 1, unicode(company, "utf-8"))
	  sheet.write(cur_row, 4, unicode(link, "utf-8"))
	  sheet.write(cur_row, 0, unicode(date, "utf-8"))

def processSheet(sheet):
	  sheet.col(3).width = 256 * 60
	  sheet.col(1).width = 256 * 15
	  sheet.col(4).width = 256 * 100
	  sheet.col(0).width = 256 * 10

def savewb(wb):
	  wbname = 'Google_Alerts_%s' % (datetime.now().strftime("%m-%d-%y")) + '.xls'
	  wbname = os.getcwd() + '/alerts_spreadsheets/' + wbname

	  print 'Processing completed...'
	  print 'Saving to ' + wbname + '...'
	  wb.save(wbname)

def createSpreadsheet():
	  wb = xlwt.Workbook()
	  sheet = wb.add_sheet("Google Alerts")
	  style = xlwt.easyxf('font: bold 1')
	  sheet.write(0, 3, 'Headline', style)
	  sheet.write(0, 1, 'Company', style)
	  sheet.write(0, 4, 'URL', style)
	  sheet.write(0, 0, 'Date', style)

	  cur_row = 1

	  for url in LA_HONDA_ALERTS_URLS:
			print 'Processing google alerts for ' + LA_HONDA_ALERTS_URLS[url] + '...'
			r = requests.get(url)
			xml = r.text
			soup = BeautifulSoup(xml)

			for title, link, date in zip(soup.findAll('title')[1:], soup.findAll('link')[1:], soup.findAll('published')):
				  title = cleanTitle(title)
				  link = cleanLink(link)
				  date = cleanDate(date)

				  writeToSheet(sheet, title, LA_HONDA_ALERTS_URLS[url], link, date, cur_row)
				  cur_row = cur_row + 1

	  processSheet(sheet)
	  savewb(wb)

USERNAME = 'sorcererdailyupdate@gmail.com'
PASSWORD = 'CrothersStoreySidDrew2015'
DAVID_EMAIL = ['david@lahondaadvisors.com']

def main():
	  createSpreadsheet()
	  wbname = os.getcwd() + '/alerts_spreadsheets/' + 'Google_Alerts_%s' % (datetime.now().strftime("%m-%d-%y")) + '.xls'
	  date = datetime.now()
	  subject = 'Google Alerts Update for ' + date.strftime('%m-%d-%Y')
	  email_body = 'Hi David,\n\nAttached are Google Alerts updates for ' + date.strftime('%m-%d-%Y') + '\n\nBest,\nSid'

	  email_digest = EmailDigestAPI(USERNAME, PASSWORD)
	  email_digest.send_mail(DAVID_EMAIL, subject, email_body, files = [wbname])

if __name__ == "__main__":
	  main()
