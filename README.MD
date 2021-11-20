               LOCAL BUCKLING ANALYSIS OF A  LATERAL ROGUE BUCKLED SUBSEA RIGID PIPELINE UNDER LOAD CONTROLLED CONDITION (LCC).

Tools used:
 -  Abaqus FEA
 -  Python
 -  xlsxwriter

Code & Standard:
- DNVGLSTF101
- DNVGLRPF110

Description:
Local buckling implies gross deformation of the cross-section of the pipeline.

This analysis is about the determination of the Local buckling utilization of a pipeline rogue buckle to ascertain if the rogue buckle meets acceptance criteria for a combined loading criteria according to DNVGLSTF101 Load Controlled Condition (LCC).  Lateral buckling is not a failure mode on its own, but if not contained can lead to local buckling, fatigue and fracture which are failure modes that can affect the integrity of the pipeline. 

The section moment, effective axial force taken from abaqus lateral buckling analysis are important input that are combined with the load effect factors to form the design bending moment (Msd) , design effective axial force (Ssd). The combined loading criteria for LCC is a combination of Msd, SSd and internal overpressure (in the case of operating or hydrotest conditions) or external overpressure (in the case of installation or shutdown) with their load effect factors. The final local buckling utilization should be less than 1 for it  to be acceptable.

The following steps were considered in the FEA analysis :
- STEP1: External Pressure ,emptyweight, RLT
- STEP2: Introduce buckle
- STEP3: Release buckle
- STEP4: Remove Residual Lay Tension and apply friction for empty case
- STEP5: Flooded condition
- STEP6: Hydrotest condition
- STEP7: Dewater
- STEP8: Apply Operating condition.
- STEP9: Shutdown condition

Relevant results are requested to be written to the Abaqus output database (.odb).
The analysis is runned and successfully completed.

POSTPROCESSING:
A python script was written that search the directory where the odb is, open the odb and extract the relevant results and then use then perform calculations to get the LCC utilization
1. Pipeline coordinate in x-direction - obtained from Abaqus
2. Section moment (SM2) - obtained from Abaqus
3. Effective Axial Force (ESF1) - obtained from Abaqus
4. Calculation of the derating values due to temperature of the yield stress and tensile stress above 50oc for C-Mn steel pipe
5. Calculation of the plastic moment capacity (Mp)  and plastic force capacity (Sp)
6.Calculation of burst pressure (Pb) and characteristic collapse pressure (Pc)
4. Calculation of the design effective axial force(Ssd) and ombination with their load effect factors
5. Calculation of the design bending moment (Msd) and ombination with their load effect factors
6. Combination with the presssure terms and their load effect factors
7. Calculation of the local buckling utilization









