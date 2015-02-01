
SELECT *
FROM(
    SELECT  P.PROCESS_ID, P.STEP_SEQ, E.PROCESS_GROUP_NAME, P.EQP_ID||'-'||P.TKIN_PREVENT_CHAMBER_ID As Tool, max(S.RECIPE_ID) As PPID ,P.TKIN_PREVENT_TYPE
    FROM    SMIMES.MI_TKIN_PREVENT P  

    JOIN  ( SELECT Line_ID, SUBSTR(Eqp_ID, 1, 6) AS Eqp_ID, NVL(SUBSTR(Eqp_ID, 8, 1), '-') AS Chamber, 
            Process_Group_Name, EQP_STATUS, EQP_TRANSN_COMMENT 
            FROM SMIMES.MI_EQUIPMENT 
            WHERE  SUBSTR(Eqp_ID,1,6) LIKE 'WFA%'   
              AND  Process_Group_Name LIKE 'CU%'   
              AND Tool_Kind IN ('EQP', 'CHAMBER')
          ) E 
         ON P.Line_ID = E.Line_ID AND P.EQP_ID = E.Eqp_ID AND P.TkIn_Prevent_Chamber_ID = E.Chamber 

    JOIN  ( SELECT Line_ID, Process_ID, Step_Seq, Recipe_ID 
            FROM SMIMES.MI_STEP
            WHERE  Step_Seq LIKE 'F%'   
              AND  Recipe_ID LIKE 'PWAR2'  
          ) S 
         ON P.Line_ID = E.Line_ID AND P.STEP_SEQ = S.STEP_SEQ AND (P.Process_ID = '-' OR P.Process_ID = S.Process_ID)

WHERE   P.LINE_ID = 'SFBX' AND P.STEP_SEQ != '-'  
GROUP BY P.EQP_ID||'-'||P.TKIN_PREVENT_CHAMBER_ID, E.EQP_STATUS, E.EQP_TRANSN_COMMENT, P.PROCESS_ID, P.STEP_SEQ, E.PROCESS_GROUP_NAME, P.TKIN_PREVENT_TYPE
    )

PIVOT
    (
    MAX(TKIN_PREVENT_TYPE) FOR TOOL IN ( tool_list )
    )