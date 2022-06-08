select  *  from GIPHitSession

group by IPCat

update GIPHitSession
set IPCat = (SELECT TOP 1 IPcat FROM GipHitIPCat WHERE IPnum BETWEEN IP_FROM AND IP_TO)
where gHitSID >0 and  gHitSID <=23006

SELECT * FROM GIPHitIPCat WHERE  3544914294 BETWEEN IP_FROM AND IP_TO

*********************�ӷ���O�έp********************************
SELECT         COUNT(*) AS Expr1, IPCat
FROM             gipHitSession
WHERE         (hitFirst BETWEEN '2004/6/1' AND '2004/6/30')
GROUP BY  IPCat

*********************���O�έp********************************
select iCtUnit, max(n.CatName), (count(*)*130) as hitNumber
 from gipHitUnit AS u LEFT JOIN CatTreeNode AS n ON u.iCtUnit=n.CtNodeID
where hitTime between '2004/1/1'  AND '2004/6/30' 
group by iCtUnit
order by count(*) DESC

*********************��Ƴ椸�έp********************************
select u.iCuItem, max(n.sTitle) as UnitName,( count(*)*130) as HitNumber
 from gipHitPage AS u LEFT JOIN CuDTGeneric AS n ON u.iCuItem=n.iCuItem
where hitTime between '2004/1/1'  AND '2004/6/30' 
group by u.iCuItem
order by count(*) DESC

********************�u��GIP�����O�έp�έp(�h��KM)*************************
select iCtUnit, max(n.CtUnitName), count(*) as hitNumber
 from gipHitUnit AS u LEFT JOIN CtUnit AS n ON u.iCtUnit=n.CtUnitID
where hitTime between '2004/1/1'  AND '2004/8/30' and u.kmcat is null
group by iCtUnit
order by count(*) DESC