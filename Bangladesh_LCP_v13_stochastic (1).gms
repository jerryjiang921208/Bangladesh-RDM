*Option LP=PATHNLP;


Option iterlim=1000000000;
Option reslim=1000000000;

Sets

g generators

t time increment of LDC

f fuel type

y years

s scenarios considered

sce_id/0*1/

map(g,f)

;


alias (y,y1);


Scalar

IncludeCO2Price /0/

IncludeEnergyEfficiency /0/
;

Scalar r discount rate /0.06/
*as per Bank's advice on climate resilient projects

       MaxCapital max gen investment in billion dollars /132.5/

       MAXCOALCAP  max GW of coal that can be built /200/
;

Parameters prob(s);
$Call GDXXRW   Bangladesh_InputsCO2.xlsx  set=s RDim=1 rng=probabilities!A2:A6 par=prob RDim=1  rng=probabilities!A2:B6
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD s
$LOAD prob
$GDXIN
;
display s, prob;
*note that if you want to include more scenarios you have to update the range in excel


Parameters GenData(g,*),  FuelPrices(f,y), FuelLimit(f,y), Duration(t), EmisFactor(f,*), CO2Price(y,*), Import_power_price(g,t) ,growth_import, cap_multiplier(y),solar_cf(t),CFmax(g,s),FOM(g,s),flood_prote_gas(y),interconnection_limit(y),include(g)
;


$Call GDXXRW   Bangladesh_InputsCO2.xlsx  set=g  RDim=1  rng=GenData!B12:B1000  par=GenData  RDim=1  CDim=1  rng=GenData!B11:aa1000 par=include RDim=1 rng=Hard_criterion!a2:b154
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD g
$LOAD GenData
$LOAD include
$GDXIN
* this is used to define the elements of set g and indicate the range of data defined at the g dimension

$Call GDXXRW  Bangladesh_InputsCO2.xlsx  set=t  RDim=1  rng=LDC!A4:A1000   set=y CDim=1 rng=LDC!j3:x3 par=Duration  RDim=1 rng=LDC!a4:b1000 par=flood_prote_gas CDim=1 rng=Resource_potential_time!b1:Q2  par=interconnection_limit CDim=1  rng=Resource_potential_time!b3:Q4
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD t
$LOAD y
$LOAD Duration
$LOAD flood_prote_gas
$LOAD interconnection_limit
$GDXIN
*this is used to define the set of time blocks to be used in the model and load the load data

$Call GDXXRW  Bangladesh_InputsCO2.xlsx  set=f  RDim=1  rng=FuelPrices!A3:A20   par=FuelPrices RDim=1 CDim=1 rng=FuelPrices!A2:R20 par=CO2Price RDim=1 CDim=1 rng=Emission!E4:F1000
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD f
$LOAD FuelPrices
$LOAD CO2Price
$GDXIN

*this is used to define the set of fuels in the model and load their prices


$Call GDXXRW  Bangladesh_InputsCO2.xlsx  set=map  RDim=1 CDim=1  rng=map!A2:U990 par=FuelLimit RDim=1 CDim=1 rng=FuelLimit!A2:AA1000 par=EmisFactor RDim=1 CDim=1 rng=Emission!A4:B1000
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD map
$LOAD FuelLimit
$LOAD EmisFactor
$GDXIN

$Call GDXXRW  Bangladesh_InputsCO2.xlsx     par=Import_power_price  rng=Interconnection_price!A1:AE6  par=cap_multiplier RDim=1 rng=Capital_multipliers!A9:B23
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD Import_power_price
$LOAD cap_multiplier
$GDXIN
*note that if you intend to change the initial year of the model, capital multiplier has to change too.
$Call GDXXRW  Bangladesh_InputsCO2.xlsx     par=solar_cf RDim=1 rng=Solar_profile!A2:b31
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD solar_cf
$GDXIN

$Call GDXXRW  Bangladesh_InputsCO2.xlsx   par=CFmax  rng=CF!A3:f201
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD CFmax
$GDXIN

$Call GDXXRW  Bangladesh_InputsCO2.xlsx   par=FOM rng=FOM!A2:f200
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD FOM
$GDXIN

Parameters block_derate(t,s);
$Call GDXXRW   Bangladesh_InputsCO2.xlsx  par=block_derate    rng=Derate_assumption!e2:i32
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD block_derate
$GDXIN
;

$ontext
parameters ldc_growth(m);
$Call GDXXRW   Bangladesh_InputsCO2.xlsx   par=ldc_growth RDim=1  rng=sce_mat!a2:B401
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD ldc_growth
$GDXIN

display ldc_growth
$offtext

parameters Adrien(sce_id,*);
$Call GDXXRW   Bangladesh_InputsCO2.xlsx   par=Adrien RDim=1  CDim=1  rng=sce_mat!a1:k3
$GDXIN  Bangladesh_InputsCO2.gdx
$LOAD Adrien
$GDXIN



display block_derate;
display FOM;
display CFmax;
display flood_prote_gas;
display cap_multiplier;
PARAMETER LDC(s,t,y)
/
$ondelim
$include demand_det_old_2016.txt
$offdelim
/;
*demand_stoch_old has the projection for LDC initially provided by the bank perturbed and the demand.txt has the one built by Elina




* Interconnection prices from adjacent markets/systems loaded

Parameter EEScalingFactor(y)
/

2016    0.89
2017    0.88
2018    0.87
2019    0.86
2020    0.85
2021    0.845
2022    0.84
2023    0.835
2024    0.83
2025    0.825
2026    0.82
2027    0.815
2028    0.81
2029    0.805
2030    0.80
/;
*2015    0.90
*2014    1.00
Parameter PSsaving(y)
/
2016    0.027
2017    0.03
2018    0.032
2019    0.035
2020    0.037
2021    0.039
2022    0.04
2023    0.041
2024    0.042
2025    0.044
2026    0.045
2027    0.046
2028    0.047
2029    0.049
2030    0.05
/;
*2014    0
*2015    0.025
*assumed based on the ratio of PS has been considered in the previous case
Parameter growth_import
/
0.02
/;

parameter VOM(g), PMIN(g), PDerated(g), CAPEX(g), Efficiency(g),  FCdom(g), FCimp(g), ng(g), Operation(g),Technology(g), Fuel_type(g);

VOM(g)                   =GenData(g,"VOM");
PDerated(g)              =GenData(g,"PDerated");
Efficiency(g)            =GenData(g,"Efficiency");
FCdom(g)                 =GenData(g,"Fuel_Cost_Dom");
FCimp(g)                 =GenData(g,"Fuel_Cost_Imp");
Operation(g)             =GenData(g,"Operation");
ng(g)                    =YES$(Operation(g) = 3);
Technology(g)            =GenData(g,"Technology");
Fuel_type(g)             =GenData(g,"Fuel");
CAPEX(g)                 =GenData(g,"ANN_Capex_per_MW WITH elevation (eu-watch)");


display CAPEX;
display FuelLimit;


Parameter VC(g,y,f,t), GenEmis(g,f);

*Variable cost of fuel $/MWh-e
VC(g,y,f,t)$(map(g,f)and Technology(g)<>6)       =FuelPrices(f,y)/(0.293071*Efficiency(g));
VC(g,y,f,t)$(map(g,f)and Technology(g)=6)= Import_power_price(g,t);
*The second expressions assigns the interconnections prices
display VC;
display include;

Parameter Initial_capacity(g) ;
Initial_capacity(g)$(GenData(g,"CapexperMW"))=0;
Initial_capacity(g)$(GenData(g,"CapexperMW")=0 and GenData(g,"StYr") <= 2015)=GenData(g,"Pderated");
Initial_capacity(g)$(GenData(g,"CapexperMW")=0 and GenData(g,"StYr") > 2015)=0;
*candidates have zero initial capacity same for already planned generators
*existing generators have the maximum capacity


*GenEmis is emission factor in tonne per MWhe
GenEmis(g,f)$map(g,f)    =(EmisFactor(f,"CO2")/(0.293071*Efficiency(g)))/1000;



Variables
Gen(s,g,f,y,t)          generation per unit per year per LDC point in MWh
Cap(s,g,y)              installed capacity in year y in MW
Build_1st_stage(g,y)    Build new capacity MW in year y
Build_2nd_stage(s,g,y)  Build new capacity MW in year y
USE(s,y,t)              unserved energy in MW
Unmetreserve(s,y)       reserve capacity shortfall
Surplus(s,y,t)          surplus power (to get around the min load constraint!)
Fuel(s,f,y)             fuel consumption in MMBTU
cost                    total system cost  in USD
Retire_1st_stage(g,y) retirement
Retire_2nd_stage(s,g,y) retirement
 ;


Positive variables gen, cap,Build_1st_stage ,Build_2nd_stage, fuel, USE, Surplus, UnmetReserve,Retire_1st_stage,Retire_2nd_stage;
*play with scenarios  (if scenario has zero probability do not include the decision variable for easier
Gen.fx(s,g,f,y,t)$(prob(s)=0)=0;
Cap.fx(s,g,y)$(prob(s)=0)=0;
USE.fx(s,y,t)$(prob(s)=0)=0;
Unmetreserve.fx(s,y)$(prob(s)=0)=0;
Surplus.fx(s,y,t)$(prob(s)=0)=0;
Fuel.fx(s,f,y)$(prob(s)=0)=0;
Retire_2nd_stage.fx(s,g,y)$(prob(s)=0)=0;
Build_2nd_stage.fx(s,g,y)$(prob(s)=0)=0;



*Generation can take positive values for the fuel that matches each of the generators
Gen.fx(s,g,f,y,t)$(NOT map(g,f))=0;

* Existing capacity i.e., anything with capex = 0 is given the opportunity to be built up to Pderated after its starting year before it is fixed to zero
Cap.up(s,g,y)$(GenData(g,"CapexperMW")=0 and GenData(g,"StYr") le Ord(y)+2015)=GenData(g,"Pderated");
Cap.fx(s,g,y)$(GenData(g,"CapexperMW")=0 and GenData(g,"StYr") > Ord(y)+2015)=0;

*No new build at existing sites- they can be built (the planned ones only in their starting year- we could relax that for later years too)
Build_1st_stage.fx(g,y)$(GenData(g,"CapexperMW")=0 and GenData(g,"StYr")-2015<>ord(y) )=0;
Build_2nd_stage.fx(s,g,y)$(GenData(g,"CapexperMW")=0 and GenData(g,"StYr")-2015<>ord(y))=0;

*for the existing ones, consider forced retirements within their horizon
*note that we do not consider rebuild since the number of variable would unnecessarily increase, the new candidates have better characteristics
*retirement at the end of the year (this adds one year if they differ by lifetime (minor for time being)
Cap.fx(s,g,y)$(GenData(g,"RtrYr") -2015< Ord(y))=0;


*by adjusting the year (for example now 17 we move the year when the uncertainty is resolved)
*we constraint the domain of the variables build_1st_stage and build_2nd_stage
Build_1st_stage.fx(g,y)$(ord(y)>17)=0;
Build_2nd_stage.fx(s,g,y)$(ord(y)<=17)=0;
Retire_1st_stage.fx(g,y)$(ord(y)>17)=0;
Retire_2nd_stage.fx(s,g,y)$(ord(y)<=17)=0;

* cannot build any new gen unit before its starting year (StYr
Build_1st_stage.fx(g,y)$(GenData(g,"CapexperMW") and GenData(g,"StYr") > Ord(y)+2015 )=0;
Build_2nd_stage.fx(s,g,y)$(GenData(g,"CapexperMW") and GenData(g,"StYr") > Ord(y)+2015 )=0;

*experiment 2 : do not provide any new import option
*Build_1st_stage.fx(g,y)$(Technology(g)=6 and ord(y)>1 )=0;
*Build_2nd_stage.fx(s,g,y)$(Technology(g)=6 and ord(y)>1 )=0;

* Peak shaving can only be implemented for limited hours every year. Initial assumption (v6) was that all first 10 blocks can use it
*but this assumption is equivalent to 3 hours every day of the year, seems a bit high. Assume again 3 hours but 30% of the year
*this is equivalent to the first 6 blocks  more or less
Gen.fx(s,"PS",f,y,t)$(ord(t)>6) = 0;
*to implement the hard criterion (I give flexibility up to 2020 for units to retire if needed
Gen.fx(s,g,f,y,t)$(include(g)=0 AND ord(y)>=5 )=0;
Cap.fx(s,g,y)$(include(g)=0 and ord(y)>=5)=0;
Build_1st_stage.fx(g,y)$(include(g)=0)=0;
Build_2nd_stage.fx(s,g,y)$(include(g)=0)=0;



Equations

Obj Objective function
Capacity(s,g,y,t)
Intermittent_capacity(s,g,y,t)

DemBal(s,y,t)


CapBal(s,g,y)              capacity balance
CapBal1st(s,g,y)           capacity balance 1st year

MaxBuild(s,g)

MinCapReserve(s,y)         minimum capacity reserve

*JointFuel(s,g,y,t)

MaxCF(s,g,y)               maximum capacity factor limit for each generator

*MinLoad(g,y,t)           minimum loading limit

FuelBal(s,f,y)             fuel balance

FuelLimitCon(s,f,y)        gas limits

CapitalConstraint(s)        total capital on new investment is constrained

CoalCapConstraint(s)        max coal capacity addition limited to MAXCOALCAP GW
Floodsafe(y,s)              maximum MW in flood-safe existing gas power plant locations that are available due to retirements
Energyefficiency(y,s)       maximum capacity by EE every year
Peakshaving(y,s)            maximum capacity by PS every year
*the coal capital constraint applies to any scenario
Intercon(y,s)               maximum capacity of the interconnection to adjacent systems
*Smooth_LNG_1(y,s)             constraint to make sure that smooth transitions from inter-annual variability in LNG use
*Smooth_LNG_2(y,s)             constraint to make sure that smooth transitions from inter-annual variability in LNG use
;

Obj ..  cost =e=
* note that depending on the year it is build, cap_multiplier helps us to include all the annualized capital
*payments it would pay up to 2030 (note that everytime the r changes we have to update the excel or I can include it in the model)
Sum((y,g), Build_1st_stage(g,y)*CAPEX(g)*cap_multiplier(y) ) + Sum((y,g,s),prob(s)* Build_2nd_stage(s,g,y)*CAPEX(g)*cap_multiplier(y) )  +
Sum((s,y),prob(s)* (1/(1+r)**(ord(y)-1))*(
Sum((g), Cap(s,g,y)*FOM(g,s))

* dispatch cost - fuel in $
+   sum((g,f)$map(g,f), Sum((t), Gen(s,g,f,y,t)*Duration(t)*VC(g,y,f,t)))


* CO2 Cost in $
+   (sum((g,f)$map(g,f), Sum((t), Gen(s,g,f,y,t)*Duration(t)*GenEmis(g,f)*CO2Price(y,"CO2Price"))))$IncludeCO2Price

* VOM cost in $
+   sum((g,f)$map(g,f), Sum((t), Gen(s,g,f,y,t)*Duration(t)*GenData(g,"VOM")))

* unserved energy cost valued at USD 500 per MWh or 50 cents per kWh
+  Sum((t), USE(s,y,t)*Duration(t)*500  )

* Surplus energy cost valued at 100$ per MWh (not sure I am correct about this)
+  Sum((t), Surplus(s,y,t)*Duration(t)*100  )


* Any reserve MW shortfall is penalized with 100K per MW per year (22nd May- before it was 500K per MW)
+ UnmetReserve(s,y)*100000

))
;


*  The capacity of the new year is equal to the new built capacity and the capacity of the previous year
CapBal(s,g,y)$(ord(y)>1 and prob(s)<>0).. Cap(s,g,y) =e= Cap(s,g,y-1)+Build_1st_stage(g,y)+Build_2nd_stage(s,g,y)-Retire_1st_stage(g,y)-Retire_2nd_stage(s,g,y);

CapBal1st(s,g,y)$(ord(y)=1 and prob(s)<>0)..  Cap(s,g,y) =e= Initial_capacity(g)+Build_1st_stage(g,y)+Build_2nd_stage(s,g,y)-Retire_1st_stage(g,y)-Retire_2nd_stage(s,g,y);

* capacity must exceed peak demand by 15% (as per Deb's advice on May 4th 2016)
*Note that for PV (which is technology 8, we do not count any contribution to the reserves
*since at the peak block it has CF=0
MinCapReserve(s,y)$(prob(s)<>0).. Sum(g$(Technology(g)<>8),Cap(s,g,y)) + UnmetReserve(s,y) =g= 1.15*Smax(t, LDC(s,t,y));

*  For each generator the total capacity built can not exceed the derated capacity
MaxBuild(s,g)$(GenData(g,"CapexperMW")and prob(s)<>0) .. Sum(y, Build_1st_stage(g,y)+Build_2nd_stage(s,g,y)) =l= GenData(g,"Pderated");


* Maximum generation on an hourly basis MWh can not exceed capacity  MW
Capacity(s,g,y,t)$(Technology(g)<>8 and prob(s)<>0)..  Sum(f$map(g,f), Gen(s,g,f,y,t)) =l= (1+block_derate(t,s)*GenData(g,"Tech_derate"))*Cap(s,g,y);
*For solar, which is the only intermittent as of version 8 of the model the output per each time block is constrained based on the solar profile
Intermittent_capacity(s,g,y,t)$(Technology(g)=8 and prob(s)<>0)..  Sum(f$map(g,f), Gen(s,g,f,y,t)) =l= (1+block_derate(t,s)*GenData(g,"Tech_derate"))*solar_cf(t)*Cap(s,g,y);

*JointFuel(s,g,y,t).. Sum(f$map(g,f), Gen(s,g,f,y,t)) =l= Cap(s,g,y);
*joint fuel seems identical to the constraint above

* Total generation  for every generator and for every year can not exceed the generation corresponding to the maximum capacity factor
MaxCF(s,g,y)$(prob(s)<>0).. Sum(f$map(g,f), Sum(t, Gen(s,g,f,y,t)*Duration(t))) =l= CFmax(g,s)*Cap(s,g,y)*8760;


*MinLoad(g,y,t)$(GenData(g,"Pderated") and GenData(g,"StYr") le Ord(y)+2013)..       Sum(f$map(g,f), Gen(g,f,y,t)) =g=  0*GenData(g,"Pmin")/GenData(g,"Pderated")*Cap(g,y);


FuelBal(s,f,y)$(prob(s)<>0).. Fuel(s,f,y) =e= Sum(g$map(g,f), Sum(t, Gen(s,g,f,y,t)*Duration(t)/(0.293071*Efficiency(g))));

* Fuel limits in MMBTU/YEAR
FuelLimitCon(s,f,y)$(Sum(y1,FuelLimit(f,y1))and prob(s)<>0).. Fuel(s,f,y) =l= FuelLimit(f,y);

DemBal(s,y,t)$(prob(s)<>0).. Sum((g,f)$map(g,f), Gen(s,g,f,y,t)) + USE(s,y,t) - Surplus(s,y,t) =e= LDC(s,t,y);

CapitalConstraint(s)$(prob(s)<>0).. Sum(y, Sum((g), (Build_1st_stage(g,y)+Build_2nd_stage(s,g,y))*(GenData(g,"CapexperMW")+GenData(g,"Flood_protection_capex (EU-WATCH)")))) =l= MaxCapital*1000;

CoalCapConstraint(s)$(prob(s)<>0).. Sum(y, Sum(g$(GenData(g,"Fuel")=4 and GenData(g,"CapexperMW")), Build_1st_stage(g,y)+Build_2nd_stage(s,g,y))) =l= MaxCOALCAP*1000;

Floodsafe(y,s)$(flood_prote_gas(y)and prob(s)<>0) .. Cap(s,"CCGT_aggregate",y)+ Cap(s,"OCGT_aggregate",y)=l=flood_prote_gas(y);

*BiomassCoalCon(g,"Biomass",y,t)$map(g,"Biomass")..  Gen(g,"Biomass",y,t) =e= 0.15*sum(f$map(g,f), Gen(g,f,y,t));
Energyefficiency(y,s)$(prob(s)<>0)..Cap(s,"EE",y)=l=IncludeEnergyEfficiency*(1-EEScalingFactor(y))*sum(t,  LDC(s,t,y)*Duration(t)/8760);

Peakshaving(y,s)$(prob(s)<>0)..Cap(s,"PS",y)=l=PSsaving(y)*Smax(t, LDC(s,t,y));
Intercon(y,s)$(interconnection_limit(y) and prob(s)<>0)..sum(g$(Technology(g)=6 and Operation(g) = 3), Cap(s,g,y))=l= interconnection_limit(y);

*Smooth_LNG_1(y,s)$(ord(y)>1 and prob(s)<>0)..Fuel(s,"LNG",y)=l=1.25*(Fuel(s,"LNG",y-1)-188);
*Smooth_LNG_2(y,s)$(ord(y)>1 and prob(s)<>0)..Fuel(s,"LNG",y)=g=0.75*Fuel(s,"LNG",y-1);

Model BangladeshLCP_stoch /all/;

Solve BangladeshLCP_stoch using lp minimizing cost;


Parameters First_stage(y,f), Second_stage(s,y,f),Genrn(s,f,y),Annual_cost(s,y),reported_cost,Served_energy(s,y),Unserved_energy(s,y),Unmet_reserve_cost(s,y),import_share(s,y),CO2_emissions(s,y),capex_cost(s),FuelCons(s,f,y),FuelConShadowPrice(s,f,y),Electricity_price(s,y,t),variable_cost(s),capital_included(s),underved_cost(s),FOM_incl(s);


First_stage(y,f) = sum(g$map(g,f), Build_1st_stage.l(g,y));
Second_stage(s,y,f) =  sum(g$map(g,f), Build_2nd_stage.l(s,g,y));
Annual_cost(s,y)$(prob(s)<>0)=
Sum((g,y1)$(ord(y1)<=ord(y)), Build_1st_stage.l(g,y1)*CAPEX(g)* 1/(1+r)**(ord(y)-1)) + Sum((g), Build_2nd_stage.l(s,g,y)*CAPEX(g)*cap_multiplier(y) )  +
 (1/(1+r)**(ord(y)-1))*(
Sum((g), Cap.l(s,g,y)*FOM(g,s))

* dispatch cost - fuel in $
+   sum((g,f)$map(g,f), Sum((t), Gen.l(s,g,f,y,t)*Duration(t)*VC(g,y,f,t)))


* CO2 Cost in $
+   (sum((g,f)$map(g,f), Sum((t), Gen.l(s,g,f,y,t)*Duration(t)*GenEmis(g,f)*CO2Price(y,"CO2Price"))))$IncludeCO2Price

* VOM cost in $
+   sum((g,f)$map(g,f), Sum((t), Gen.l(s,g,f,y,t)*Duration(t)*GenData(g,"VOM"))) );

reported_cost=sum(s, prob(s)*sum(y,annual_cost(s,y))/1000000);
capex_cost(s)$(prob(s)<>0)=sum((g,y),(Build_1st_stage.l(g,y)+Build_2nd_stage.l(s,g,y))*(GenData(g,"CapexperMW")+GenData(g,"Flood_protection_capex (EU-WATCH)")));
*note that capex is not discounted and should be in million dollars
variable_cost(s)$(prob(s)<>0)=sum((y,g,f,t)$map(g,f), (1/(1+r)**(ord(y)-1))*Gen.l(s,g,f,y,t)*Duration(t)*(VC(g,y,f,t)+GenData(g,"VOM")));
capital_included(s)$(prob(s)<>0)= Sum((y,g),( Build_1st_stage.l(g,y)*CAPEX(g)*cap_multiplier(y)))  + Sum((y,g), Build_2nd_stage.l(s,g,y)*CAPEX(g)*cap_multiplier(y) );
underved_cost(s)$(prob(s)<>0)=   Sum((y,t),1/(1+r)**(ord(y)-1)*( USE.l(s,y,t)*Duration(t)*500  +   Surplus.l(s,y,t)*Duration(t)*100  ))+sum(y,1/(1+r)**(ord(y)-1)*UnmetReserve.l(s,y)*100000);
FOM_incl(s)$(prob(s)<>0)=Sum((g,y), 1/(1+r)**(ord(y)-1)*Cap.l(s,g,y)*FOM(g,s));

Served_energy(s,y)=sum((t),(LDC(s,t,y)-USE.l(s,y,t))*Duration(t));
Unserved_energy(s,y)=  sum((t),(USE.l(s,y,t))*Duration(t));
Unmet_reserve_cost(s,y)=UnmetReserve.l(s,y)*100000*(1/(1+r)**(ord(y)-1));

Genrn(s,f,y) = Sum(g$map(g,f), Sum(t, Gen.l(s,g,f,y,t)*Duration(t)))/1000;
import_share(s,y)$(prob(s)<>0)= Sum((t,g,f)$(Technology(g)=6),Gen.l(s,g,f,y,t)*Duration(t)/1000 )/SUM(f, Genrn(s,f,y));
CO2_emissions(s,y)=(sum((g,f)$map(g,f), Sum((t), Gen.l(s,g,f,y,t)*Duration(t)*GenEmis(g,f))))/1e6;


Electricity_price(s,y,t)$(prob(s)<>0)= DemBal.m(s,y,t)/Duration(t)*(1+r)**(ord(y)-1);
FuelCons(s,f,y)$(prob(s)<>0) = Fuel.l(s,f,y)/1e6;
FuelConShadowPrice(s,f,y)$(prob(s)<>0) = FuelLimitCon.m(s,f,y)*(1+r)**(ord(y)-1);



display     CO2_emissions;
EXECUTE 'GDXXRW Summary.gdx rng=USE!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Unmet_reserve_cost!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Sheet2!A1:AA100 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Gen_By_Fuel!A1:AA100 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Summary!A1:AA100 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Gen_Dispatch!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Capacity!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=build_1st_stage!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=build_2nd_stage!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=1st_Tech!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=2nd_Tech!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Retirement_1st_stage!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Retirement_2nd_stage!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Cost!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=SE!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Fuel_Mill_MMBTU!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Fuel_Limit_Price!A1:AA1000 clear o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx rng=Elec_price!A1:AA1000 clear o=BDResults.xls';

EXECUTE_UNLOAD'Summary' reported_cost,import_share,CO2_emissions,GenData, Cap.l,Build_1st_stage.l,Build_2nd_stage.l,First_stage, Second_stage,Genrn, Gen.l,Annual_cost,Served_energy,Unserved_energy,Unmet_reserve_cost,capex_cost,Retire_1st_stage.l,Retire_2nd_stage.l,Electricity_price,FuelCons,FuelConShadowPrice,variable_cost,capital_included,underved_cost,FOM_incl;
EXECUTE 'GDXXRW Summary.gdx par=GenData rng=Sheet2!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Genrn rng=Gen_By_Fuel!A5 o=BDResults.xls';
*in the summary spreadsheet you can see Deb;s requested metrics
EXECUTE 'GDXXRW Summary.gdx par=reported_cost rng=Summary!B1 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=variable_cost rng=Summary!B2 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=capital_included rng=Summary!B4 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=underved_cost rng=Summary!B6 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=FOM_incl  rng=Summary!B8 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=capex_cost rng=Summary!B16 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=import_share rng=Summary!B18 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=CO2_emissions rng=Summary!B21 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Unserved_energy rng=Summary!B24 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx var=Gen.l rng=Gen_Dispatch!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx var=Cap.l rng=Capacity!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx var=Build_1st_stage.l rng=build_1st_stage!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx var=Build_2nd_stage.l rng=build_2nd_stage!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=First_stage rng=1st_Tech!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Second_stage rng=2nd_Tech!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx var=Retire_1st_stage.l  rng=Retirement_1st_stage!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx var=Retire_2nd_stage.l  rng=Retirement_2nd_stage!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Annual_cost rng=Cost!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Served_energy rng=SE!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Unserved_energy rng=USE!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Unmet_reserve_cost rng=Unmet_reserve_cost!A5 o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=FuelCons           rng=Fuel_Mill_MMBTU!A5  o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=FuelConShadowPrice rng=Fuel_Limit_Price!A5  o=BDResults.xls';
EXECUTE 'GDXXRW Summary.gdx par=Electricity_price  rng=Elec_price!A5 o=BDResults.xls';


