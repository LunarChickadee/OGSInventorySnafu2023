___ PROCEDURE .Initialize ______________________________________________________
forcesynchronize

selectall
field Item 
sortup

;openfile "NewSupplier"
;openfile "43OGSpurchasing"

if folderpath(dbinfo("folder","")) CONTAINS "sarah" and info("databasename") contains "linked"
    Field "IDNumber"
    SortUp
    SelectDuplicates ""
    if info("empty")
        field Item 
        sortup
    else
        message "Oh no, there are duplicate IDs!"
    endif
endif
___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE .add-it-up-ordering ______________________________________________
global orderingaddup, soldwta, soldwtb, soldwtc, availwt, reord, soldwtatotal, soldwtbtotal, soldwtctotal, availwttotal, reordtotal

;;debug

arrayselectedbuild orderingaddup,¶, "45ogscomments.linked",
str(«42sold»)+¬+str(«43sold»)+¬+str(«45sold»)+¬+str(«45available»)+¬+str(«min»)+¬+str(«unit conversion»)+¬+str(«42sold»*«unit conversion»)
+¬+str(«43sold»*«unit conversion»)+¬+str(«45sold»*«unit conversion»)+¬+str(«45available»*«unit conversion»)+¬+str(«min»*«unit conversion»)

;;displaydata orderingaddup

arrayfilter orderingaddup, soldwta,¶,extract(import(),¬,7)
arrayfilter orderingaddup, soldwtb,¶,extract(import(),¬,8)
arrayfilter orderingaddup, soldwtc,¶,extract(import(),¬,9)
arrayfilter orderingaddup, availwt,¶,extract(import(),¬,10)
arrayfilter orderingaddup, reord,¶,extract(import(),¬,11)

;;displaydata reord

arraynumerictotal soldwta,¶, soldwtatotal
arraynumerictotal soldwtb,¶, soldwtbtotal
arraynumerictotal soldwtc,¶, soldwtctotal
arraynumerictotal availwt,¶, availwttotal
arraynumerictotal reord,¶, reordtotal
;;message reordtotal

if info("windowtype") = 15
    drawobjects
endif
___ ENDPROCEDURE .add-it-up-ordering ___________________________________________

___ PROCEDURE .createrepack ____________________________________________________
global vItem, vList, all_items, num_items, last_item, vRepacknum, vWarehouseID, vWarehouseitem1, vWarehouseitem2, vWarehouseitem3, 
vWarehouseitem4, vWarehouseArray, vWeight, vAmounttomake, vRepackType, wCommentsLinked
    vList = ""
    vItem = Item
    vWarehouseID = ""
    vWeight = ""
    vRepackType = ""

/*****If "Type of Repack" field is empty macro will stop and prompt a message *****/

if 
    «Type of Repack» = ""
    message "Type of Repack Field is empty"
    stop
endif

/*****Force Syncs CL in case it hasn't been done recently*****/

synchronize  

/*****This macro first closes the repack DB if it is still open *****/

call .closerepackfile

   
/***Make sure datasheet is open ***/
   
   gosheet
   wCommentsLinked=info("windowname") 
   
 
vWarehouseitem1 = ""
vWarehouseitem2 = ""
vWarehouseitem3 = ""
vWarehouseitem4 = ""
vWarehouseArray = ""


//****Populates vWeight for mixes

vWeight  = ActWt

//****Opens Staff file to populate repack and recorder names.

openfile "Staff"

//****Builds an array to populate field in repack DB that shows what repack items to use.
   
openfile "45ogscomments.warehouse"
gosheet
window wCommentsLinked
field «Repack Supplies»

arraylinebuild vWarehouseID,¶,"", «Repack Supplies»
vWarehouseitem1 = extract(vWarehouseID, ¶, 1)   
vWarehouseitem2 = extract(vWarehouseID, ¶, 2)
vWarehouseitem3 = extract(vWarehouseID, ¶, 3)
vWarehouseitem4 = extract(vWarehouseID, ¶, 4)
window "45ogscomments.warehouse"
    select str(IDNumber) contains vWarehouseitem1
    vWarehouseArray=«Item»+¬+«Description»+¶
    select str(IDNumber) contains vWarehouseitem2
    vWarehouseArray=vWarehouseArray+«Item»+¬+«Description»+¶
    select str(IDNumber) contains vWarehouseitem3
    vWarehouseArray=vWarehouseArray+«Item»+¬+«Description»+¶
    select str(IDNumber) contains vWarehouseitem4
    vWarehouseArray=vWarehouseArray+«Item»+¬+«Description»
    arraydeduplicate vWarehouseArray,vWarehouseArray,¶

window wCommentsLinked

//****Regardless of which repack is needed OGSRepack will need to be
//****opened, a new record added for item being made, and the perm
//****ID filled into repack DB by calling the permID macro
//****Go back to CL to decide which form needs to be opened based on
//****what is displayed in the Form field

case «Type of Repack» ="Seed"
    vRepackType = «Type of Repack»
    call .lotnumbers
    openfile "45ogsrepack"
    gosheet 
       
    /***Creating new Repack Number***/
    field RepackNumber
    sortup
    lastrecord
    addrecord
    arraybuild all_items,¶,"",RepackNumber
    arraynumericsort all_items,all_items,¶
    num_items = arraysize(all_items,¶)
    last_item = val(array(all_items,num_items,¶))
    clipboard() = last_item + 1
    RepackNumber=clipboard()
    
    field ItemNumMade1
    ItemNumMade1 = vItem
    «RepackSupply1» = val(vWarehouseitem1)
    «RepackSupply2» = val(vWarehouseitem2)
    «RepackSupply3» = val(vWarehouseitem3)
    «Repack Type» = vRepackType
    call .permID 

        
        openform "Seed Repack Form"  
    
case «Type of Repack» = "Seed/Mix"
       YesNo "Are you making a Repack? Click No if you are making a mix."
            if clipboard() contains "Yes"
                vRepackType = "Seed"
                call .lotnumbers
                openfile "45ogsrepack"
                gosheet    
                /***Creating new Repack Number***/
                field RepackNumber
                sortup
                lastrecord
                addrecord
                arraybuild all_items,¶,"",RepackNumber
                arraynumericsort all_items,all_items,¶
                num_items = arraysize(all_items,¶)
                last_item = val(array(all_items,num_items,¶))
                clipboard() = last_item + 1
                RepackNumber=clipboard()
                
                ItemNumMade1 = vItem
                «RepackSupply1» = val(vWarehouseitem1)
                «RepackSupply2» = val(vWarehouseitem2)
                «RepackSupply3» = val(vWarehouseitem3)
                «Repack Type» = vRepackType
                call .permID 
               
                openform "Seed Repack Form" 
            else
                GetscrapOK "How many do you need to make?"
                vAmounttomake=val(clipboard())
                call .mixrepack
                openfile "45ogsrepack"
                gosheet    
                
                /***Creating new Repack Number***/
                field RepackNumber
                sortup
                lastrecord
                addrecord
                arraybuild all_items,¶,"",RepackNumber
                arraynumericsort all_items,all_items,¶
                num_items = arraysize(all_items,¶)
                last_item = val(array(all_items,num_items,¶))
                clipboard() = last_item + 1
                RepackNumber=clipboard()
                
                TotalWt = vAmounttomake * vWeight
                AmountToMake = vAmounttomake
                
                openform "Mix Repack Form"    
                «Repack Type» = vRepackType
                
            endif
case «Type of Repack» ="Kit"
    vRepackType = «Type of Repack»
   GetscrapOK "How many do you need to make?"
                vAmounttomake=val(clipboard())
                call .kitrepack
                openfile "45ogsrepack"
                gosheet    
                /***Creating new Repack Number***/
                field RepackNumber
                sortup
                lastrecord
                addrecord
                arraybuild all_items,¶,"",RepackNumber
                arraynumericsort all_items,all_items,¶
                num_items = arraysize(all_items,¶)
                last_item = val(array(all_items,num_items,¶))
                clipboard() = last_item + 1
                RepackNumber=clipboard()
                
                «Repack Type» = vRepackType
                ItemNumMade1 = vItem
                «RepackSupply1» = val(vWarehouseitem1)
                «RepackSupply2» = val(vWarehouseitem2)
                «RepackSupply3» = val(vWarehouseitem3)
                call .permID
                AmountToMake = vAmounttomake
                openform "Kit Repack Form"
                
case «Type of Repack» ="Item"
    vRepackType = «Type of Repack»
    openfile "45ogsrepack"
    gosheet    
    
    /***Creating new Repack Number***/
    field RepackNumber
    sortup
    lastrecord
    addrecord
    arraybuild all_items,¶,"",RepackNumber
    arraynumericsort all_items,all_items,¶
    num_items = arraysize(all_items,¶)
    last_item = val(array(all_items,num_items,¶))
    clipboard() = last_item + 1
    RepackNumber=clipboard()
    
    «Repack Type» = vRepackType
    ItemNumMade1 = vItem
    «RepackSupply1» = val(vWarehouseitem1)
    «RepackSupply2» = val(vWarehouseitem2)
    «RepackSupply3» = val(vWarehouseitem3)
    call .permID 
    openform "Item Repack Form"
case «Type of Repack» = "Item/Mix"
       YesNo "Are you making a Repack? Click No if you are making a mix."
            if clipboard() contains "Yes"
                vRepackType = "Item"
                openfile "45ogsrepack"
                gosheet
                
                /***Creating new Repack Number***/
                field RepackNumber
                sortup
                lastrecord
                addrecord
                arraybuild all_items,¶,"",RepackNumber
                arraynumericsort all_items,all_items,¶
                num_items = arraysize(all_items,¶)
                last_item = val(array(all_items,num_items,¶))
                clipboard() = last_item + 1
                RepackNumber=clipboard()
                
                 «Repack Type» = vRepackType
                ItemNumMade1 = vItem
                «RepackSupply1» = val(vWarehouseitem1)
                «RepackSupply2» = val(vWarehouseitem2)
                «RepackSupply3» = val(vWarehouseitem3)
                call .permID
               openform "Item Repack Form"  
            else
                vRepackType = "Mix"
                GetscrapOK "How many do you need to make?"
                vAmounttomake=val(clipboard())
                call .mixrepack
                openfile "45ogsrepack"
                gosheet    
                
                /***Creating new Repack Number***/
                field RepackNumber
                sortup
                lastrecord
                addrecord
                arraybuild all_items,¶,"",RepackNumber
                arraynumericsort all_items,all_items,¶
                num_items = arraysize(all_items,¶)
                last_item = val(array(all_items,num_items,¶))
                clipboard() = last_item + 1
                RepackNumber=clipboard()
                
                «Repack Type» = vRepackType
                    ItemNumMade1 = vItem
                «RepackSupply1» = val(vWarehouseitem1)
                «RepackSupply2» = val(vWarehouseitem2)
                «RepackSupply3» = val(vWarehouseitem3)
                call .permID
                TotalWt = vAmounttomake * vWeight
                AmountToMake = vAmounttomake
                openform "Mix Repack Form"
            endif
case «Type of Repack» ="Mix"
    vRepackType = «Type of Repack»
 GetscrapOK "How many do you need to make?"
                vAmounttomake=val(clipboard())
                call .mixrepack
                openfile "45ogsrepack"
                gosheet    
                
                /***Creating new Repack Number***/
                field RepackNumber
                sortup
                lastrecord
                addrecord
                arraybuild all_items,¶,"",RepackNumber
                arraynumericsort all_items,all_items,¶
                num_items = arraysize(all_items,¶)
                last_item = val(array(all_items,num_items,¶))
                clipboard() = last_item + 1
                RepackNumber=clipboard()

                 «Repack Type» = vRepackType
                    ItemNumMade1 = vItem
                «RepackSupply1» = val(vWarehouseitem1)
                «RepackSupply2» = val(vWarehouseitem2)
                «RepackSupply3» = val(vWarehouseitem3)
                call .permID
                TotalWt = vAmounttomake * vWeight
                AmountToMake = vAmounttomake
                openform "Mix Repack Form"
                
case «Type of Repack» = "Cannalot"
     GetscrapOK "How many do you need to make?"
                vAmounttomake=val(clipboard())
                call .mixrepack
                openfile "45ogsrepack"
                gosheet    
                
                /***Creating new Repack Number***/
                field RepackNumber
                sortup
                lastrecord
                addrecord
                arraybuild all_items,¶,"",RepackNumber
                arraynumericsort all_items,all_items,¶
                num_items = arraysize(all_items,¶)
                last_item = val(array(all_items,num_items,¶))
                clipboard() = last_item + 1
                RepackNumber=clipboard()

                 «Repack Type» = vRepackType
                    ItemNumMade1 = vItem
                «RepackSupply1» = val(vWarehouseitem1)
                «RepackSupply2» = val(vWarehouseitem2)
                «RepackSupply3» = val(vWarehouseitem3)
                call .permID
                TotalWt = vAmounttomake * vWeight
                AmountToMake = vAmounttomake
                openform "Cannalot Mix Form"
endcase





___ ENDPROCEDURE .createrepack _________________________________________________

___ PROCEDURE .clear_summaries _________________________________________________
RemoveSummaries 7
___ ENDPROCEDURE .clear_summaries ______________________________________________

___ PROCEDURE .closerepackfile _________________________________________________
global vWindowlist

vWindowlist = ""

vWindowlist = info("windows")

/* This code below determines if the Repack database is already open and closes
that file if it is. */

if
    vWindowlist contains "45ogsrepack"
        openfile "45ogsrepack"
        save
        closefile
        nop
endif

vWindowlist = ""
___ ENDPROCEDURE .closerepackfile ______________________________________________

___ PROCEDURE .delete __________________________________________________________
displaydata Description +¬ + ?(UnitNumber = "", "", UnitNumber) + " " + ?(UnitName = "", "", UnitName)
___ ENDPROCEDURE .delete _______________________________________________________

___ PROCEDURE .depotorders _____________________________________________________
global vItem

vItem = Item

////Select items in tally that contain vItem

Openfile "45ogstally"

Select
    Order contains str(vItem) and ShipCode = "J" and Status <> "Com"

___ ENDPROCEDURE .depotorders __________________________________________________

___ PROCEDURE .dropship ________________________________________________________
global ordersize, export, given, pon, suppinfo, intorder, progress, poitemarray
suppinfo=""
intorder=""
poitemarray=""
local waswindow, RecCount, ItemCount, POL#
waswindow=info("windowname")

yesno "you sure you're ready to place order?"

if val(forordering["0-9",1])=0
    message "Try adding a few items to the order first"
    stop 
endif
if clipboard() contains "n"
    window waswindow
    stop
endif

openfile "NewSupplier"
find SupplierIDNumber=val(extract(whichsupplier, "-", 1))
arraylinebuild suppinfo, ¶, "NewSupplier", str(SupplierIDNumber)+¬+Supplier+¬+¬+¬+Contact+¬+Fax+¬+Email+¬+Phone+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")

window waswindow
forordering=arraystrip(forordering,¶)
arrayfilter forordering, intorder, ¶, extract(extract(forordering, ¶, seq()),¬, 1)+¬+extract(extract(forordering, ¶, seq()),¬, 2)+¬+
    extract(extract(forordering, ¶, seq()),¬, 3)+¬+extract(extract(forordering, ¶, seq()),¬, 4)+¬+
    extract(extract(forordering, ¶, seq()),¬, 5)+¬+extract(extract(forordering, ¶, seq()),¬, 6)+¬+
    extract(extract(forordering, ¶, seq()),¬, 7)+¬+suppinfo+¬+¬+¬+extract(extract(forordering, ¶, seq()),¬, 8)


arraysort intorder,intorder,¶


export=intorder


ordersize=extract(extract(intorder,¶, 1),¬,5)


openfile "45ogsdropship"
Synchronize
goform "Entry"
selectall
field PO
SortUp
RecCount=info("Records")

openfile "+@export"

ItemCount=info("Records")-RecCount


if ItemCount>1
    field PO
    Maximum
    Copy
    RemoveSummaries 7
    clipboard()=val(clipboard()+1)
    LastRecord
    Paste
        loop
            uprecord
            Paste
        until ItemCount-1
    copycell
    ;message clipboard()
    select PO=val(clipboard())
    field «PO Line No»
    FormulaFill seq()
    field Date
    Date=today()
    field OrderNo
    formulafill vvoice
else
    field PO
    Maximum
    Copy
    RemoveSummaries 7
    clipboard()=val(clipboard()+1)
    LastRecord
    Paste
    field «PO Line No»
    «PO Line No»=1
    field Date
    Date=today()
    OrderNo=vvoice
endif
selectall
lastrecord
if ItemCount>1
    loop
        uprecord
    until ItemCount-1
endif

lastrecord
field PO
COPY
select PO=val(clipboard())
call Gather

;message "hello"

debug
;message "history"
window waswindow
select Item = lookupselected("45ogsdropship","Item",Item,"Item","",0)
;message "last"
if info("empty")
    goto warehouse
endif

field «Purchased»
formulafill «Purchased»+lookupselected("45ogsdropship",Item,Item,"Qty",0,0)

warehouse:
openfile "45ogscomments.warehouse"
select Item=lookupselected("45ogsdropship","Item",Item,"Item","",0)

if info("empty")
else
    field «Purchased»
    formulafill «Purchased»+lookupselected("45ogsdropship",Item,Item,"Qty",0,0)
endif

window "45ogsdropship:Entry"
lastrecord

___ ENDPROCEDURE .dropship _____________________________________________________

___ PROCEDURE .export __________________________________________________________
global ordersize, vexport, suppinfo, intorder, progress, poitemarray
suppinfo=""
intorder=""
poitemarray=""
ordersize = 0
vexport = ""


local WindowCL, RecCount, ItemCount, POL#
WindowCL=info("windowname")


/*
******Adds PO information into the purchasing DB
*/


yesno "you sure you're ready to place order?"
//if forordering variable is empty this message will occur and stop the macro
if val(forordering["0-9",1])=0
   message "Try adding a few items to the order first"
   stop 
endif
//if user clicks "no" the macro will stop
if clipboard() contains "n"
   window WindowCL
   stop
//if yes is clicked the macro proceeeds
endif

debug

//open newsupplier database
openfile "NewSupplier"

//find the supplierIDnumber that equals the whichsupplier variable
find SupplierIDNumber=val(extract(whichsupplier, "-", 1))

//building an array called SUPPINFO which contains supplier contact information
arraylinebuild suppinfo, ¶, "NewSupplier", str(SupplierIDNumber)+¬+Supplier+¬+¬+¬+Contact+¬+Fax+¬+Email+¬+Phone+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")

window WindowCL

//strips the forordering array (removes blank elements)

forordering=arraystrip(forordering,¶)
arrayfilter forordering, intorder, ¶, 
    extract(extract(forordering, ¶, seq()),¬, 1)+¬+
    extract(extract(forordering, ¶, seq()),¬, 2)+¬+
    extract(extract(forordering, ¶, seq()),¬, 3)+¬+
    extract(extract(forordering, ¶, seq()),¬, 4)+¬+ 
    extract(extract(forordering, ¶, seq()),¬, 5)+¬+
    extract(extract(forordering, ¶, seq()),¬, 6)+¬+¬+
    suppinfo+¬+¬+¬+
    extract(extract(forordering, ¶, seq()),¬, 7)


//alphabetizes array elements

arraysort intorder,intorder,¶

clipboard()=intorder

vexport=intorder

//Price

ordersize=extract(extract(intorder,¶, 1),¬,5)


openfile "45OGSpurchasing"
Synchronize

//goform "Entry"

selectall
field PO
SortUp
RecCount=info("Records")

//appending data to purchasing DB from the export variable

openfile "+@vexport"

ItemCount=info("Records")-RecCount


if ItemCount>1
   field PO
   Maximum
   Copy
   RemoveSummaries 7
   clipboard()=val(clipboard()+1)
   LastRecord
   Paste
       loop
           uprecord
           Paste
       until ItemCount-1
   copycell
   ;message clipboard()
   select PO=val(clipboard())
   field «PO Line No»
   FormulaFill seq()
   field Date
   Date=today()
else
   field PO
   Maximum
   Copy
   RemoveSummaries 7
   clipboard()=val(clipboard()+1)
   LastRecord
   Paste
   field «PO Line No»
   «PO Line No»=1
   field Date
   Date=today()
endif
selectall
lastrecord
if ItemCount>1
   loop
       uprecord
   until ItemCount-1
endif

lastrecord
field PO
COPY
select PO=val(clipboard())
call Gather

window WindowCL
select Item = lookupselected("45OGSpurchasing","Item",Item,"Item","",0)
;message "last"
if info("empty")
   goto warehouse
endif

field «On Order»
formulafill «On Order»+lookupselected("45OGSpurchasing",Item,Item,"Qty",0,0)

warehouse:
openfile "45ogscomments.warehouse"
select Item=lookupselected("45OGSpurchasing","Item",Item,"Item","",0)

if info("empty")
else
   field «On Order»
   formulafill «On Order»+lookupselected("45OGSpurchasing",Item,Item,"Qty",0,0)
endif

window "45OGSpurchasing:Entry"
firstrecord


___ ENDPROCEDURE .export _______________________________________________________

___ PROCEDURE .finditem ________________________________________________________
local finditem
GetscrapOK "What's the Item# ?"

case
    length(clipboard()) = 4
    finditem=val(clipboard())
    Select «parent_code»=finditem
case
     length(clipboard()) = 6
    finditem=clipboard()
    Select Item=finditem
endcase



___ ENDPROCEDURE .finditem _____________________________________________________

___ PROCEDURE .get_summaries ___________________________________________________
Select «can be added up» contains "Y"

Field «parent_code»
GroupUp

Field «can be added up»
Maximum

Field TotalPoundsToSell
Maximum

SelectAll
___ ENDPROCEDURE .get_summaries ________________________________________________

___ PROCEDURE .huh _____________________________________________________________
global test
test = ""

arrayselectedbuild test, ¶,"",
rep(chr(32), 20-length(str(«SupplierID»)))+str(«SupplierID»)

displaydata test
___ ENDPROCEDURE .huh __________________________________________________________

___ PROCEDURE .kitrepack _______________________________________________________
///This Macro creates variables based on the kit that will be made. Each ingredient receives its own variable along with the number of ingredients needed for that kit.
///The variables are then used in the kit repack form in the repack database.

global vIngredient1, vIngredient1,vIngredient2,vIngredient3,vIngredient4,vIngredient5,vIngredient6,vIngredient7,vIngredient8,vIngredient9,vIngredient10,vIngredient11,
vIngredient12,vIngredient13,vIngredient14,vIngredient15,vIngredient16,vIngredient17, vPercentage1, vPercentage2, vPercentage3, vPercentage4, vPercentage5, vPercentage6, vPercentage7, vPercentage8, 
vPercentage9, vPercentage10, vPercentage11, vPercentage12, vPercentage13, vPercentage14

openfile "Kits"

select str(«Kit Item») contains vItem
debug
vIngredient1 = «Ingredient 1 Item»
message vIngredient1
vPercentage1 = «Amt of Ingredient 1»
vIngredient2 = «Ingredient 2 Item»
vPercentage2 = «Amt of Ingredient 2»
vIngredient3 = «Ingredient 3 Item»
vPercentage3 = «Amt of Ingredient 3»
vIngredient4 = «Ingredient 4 Item»
vPercentage4 = «Amt of Ingredient 4»
vIngredient5 = «Ingredient 5 Item»
vPercentage5 = «Amt of Ingredient 5»
vIngredient6 = «Ingredient 6 Item»
vPercentage6 = «Amt of Ingredient 6»
vIngredient7 = «Ingredient 7 Item»
vPercentage7 = «Amt of Ingredient 7»
vIngredient8 = «Ingredient 8 Item»
vPercentage8 = «Amt of Ingredient 8»
vIngredient9 = «Ingredient 9 Item»
vPercentage9 = «Amt of Ingredient 9»
vIngredient10 = «Ingredient 10 Item»
vPercentage10 = «Amt of Ingredient 10»
vIngredient11= «Ingredient 11 Item»
vPercentage11 = «Amt of Ingredient 11»
vIngredient12= «Ingredient 12 Item»
vPercentage12 = «Amt of Ingredient 12»
vIngredient13 = «Ingredient 13 Item»
vPercentage13= «Amt of Ingredient 13»
vIngredient14 = «Ingredient 14 Item»
vPercentage14 = «Amt of Ingredient 14»



___ ENDPROCEDURE .kitrepack ____________________________________________________

___ PROCEDURE .limit ___________________________________________________________
if
val(length(Description))>20
message "Description is too long"
endif
___ ENDPROCEDURE .limit ________________________________________________________

___ PROCEDURE .LiveQuery _______________________________________________________
fileglobal liveQuery, queryResult

liveclairvoyance liveQuery, 
    queryResult, 
    ¶, 
    "TheList", 
    "NewSupplier", 
    Supplier, 
    "", 
    str(SupplierIDNumber)+¬+Supplier,
    20,
    0,""





//liveclairvoyance input, outputlist, separator, listobject, database, query, compare, template, ticks, max, options

/*


liveclairvoyance l
    iveQuery, 
    queryResult, 
    ¶, 
    "TheList", 
    "NewSupplier", 
    Supplier, 
    "", 
    str(SupplierIDNumber)+¬+Supplier,
    20,
    0,""
    
    */
___ ENDPROCEDURE .LiveQuery ____________________________________________________

___ PROCEDURE .lotnumbers ______________________________________________________
/* This macro creates an array of all lot numbers for seed repacks and displays the array in
the seed repack form in the repack DB SB*/
global vLotNums 

vLotNums = ""

openfile "45CVCROPGERM"

select BaseNumber = val(left(vItem,4))

arrayselectedbuild vLotNums, ¶, "", «Supplier lot#» + ¬ + "(" + «lot#» + ")"
arraydeduplicate vLotNums, vLotNums, ¶ 

closefile

window wCommentsLinked
___ ENDPROCEDURE .lotnumbers ___________________________________________________

___ PROCEDURE .makeadjustment __________________________________________________
global vItem, wCL
vItem = Item
wCL = info("windowname")
synchronize
openfile "45Inventory Adjustment"
field Item
addrecord
Item = vItem
call .permid

message "Please confirm current 'In House' number of this form before adjusting inventory"

___ ENDPROCEDURE .makeadjustment _______________________________________________

___ PROCEDURE .mixrepack _______________________________________________________
global vIngredient1, vIngredient1,vIngredient2,vIngredient3,vIngredient4,vIngredient5,vIngredient6,vIngredient7,vIngredient8,vIngredient9,vIngredient10,vIngredient11,
vIngredient12,vIngredient13,vIngredient14,vIngredient15,vIngredient16,vIngredient17, vPercentage1, vPercentage2, vPercentage3, vPercentage4, vPercentage5, vPercentage6, vPercentage7, vPercentage8, 
vPercentage9, vPercentage10, vPercentage11, vPercentage12, vPercentage13, vPercentage14, vPercentage15, vPercentage16, vPercentage17

vIngredient1 = ""
vPercentage1 = ""
vIngredient2 = ""
vPercentage2 = ""
vIngredient3 = ""
vPercentage3 = ""
vIngredient4 = ""
vPercentage4 = ""
vIngredient5 =  ""
vPercentage5 = ""
vIngredient6 = ""
vPercentage6 = ""
vIngredient7 = ""
vPercentage7 = ""
vIngredient8 = ""
vPercentage8 = ""
vIngredient9 = ""
vPercentage9 = ""
vIngredient10 = ""
vPercentage10 = ""
vIngredient11= ""
vPercentage11 = ""
vIngredient12= ""
vPercentage12 = ""
vIngredient13 = ""
vPercentage13= ""
vIngredient14 = ""
vPercentage14 = ""
vIngredient15 = ""
vPercentage15 = ""
vIngredient16 = ""
vPercentage16 = ""
vIngredient17 = ""

openfile "Mixes"

select str(«Mix Parent Code») contains left(vItem, 4)

vIngredient1 = «ItemIngredient1»
vPercentage1 = «IngredientPercentage1»
vIngredient2 = «ItemIngredient2»
vPercentage2 = «IngredientPercentage2»
vIngredient3 = «ItemIngredient3»
vPercentage3 = «IngredientPercentage3»
vIngredient4 = «ItemIngredient4»
vPercentage4 = «IngredientPercentage4»
vIngredient5 = «ItemIngredient5»
vPercentage5 = «IngredientPercentage5»
vIngredient6 = «ItemIngredient6»
vPercentage6 = «IngredientPercentage6»
vIngredient7 = «ItemIngredient7»
vPercentage7 = «IngredientPercentage7»
vIngredient8 = «ItemIngredient8»
vPercentage8 = «IngredientPercentage8»
vIngredient9 = «ItemIngredient9»
vPercentage9 = «IngredientPercentage9»
vIngredient10 = «ItemIngredient10»
vPercentage10 = «IngredientPercentage10»
vIngredient11= «ItemIngredient11»
vPercentage11 = «IngredientPercentage11»
vIngredient12= «ItemIngredient12»
vPercentage12 = «IngredientPercentage12»
vIngredient13 = «ItemIngredient13»
vPercentage13= «IngredientPercentage13»
vIngredient14 = «ItemIngredient14»
vPercentage14 = «IngredientPercentage14»
vIngredient15 = «ItemIngredient15»
vPercentage15 = «IngredientPercentage15»
vIngredient16 = «ItemIngredient16»
vPercentage16 = «IngredientPercentage16»
vIngredient17 = «ItemIngredient17»
vPercentage17 = «IngredientPercentage17»

___ ENDPROCEDURE .mixrepack ____________________________________________________

___ PROCEDURE .permID __________________________________________________________
    Description = info("44ogscomments.linked", "Item", «Item», "Description","",0)
    Available = info("44ogscomments.linked", "Item", «Item», "44available","",0)
    In House = info("44ogscomments.linked", "Item", «Item», "44inhouse","",0)
___ ENDPROCEDURE .permID _______________________________________________________

___ PROCEDURE .placeorder ______________________________________________________
//chr(32) is a space
global orderline, orderquantity, test2, intorder, ordersize
ordersize=""

;;message seconditemlist
;;message itemlist
intorder=""

find Item=extract(itemlist," ",1)
if info("found")
    getscrap "How many "+Description+chr(32)+str(«UnitNumber»)+chr(32)+str(«UnitName»)+"?"
    orderquantity=clipboard()
        if val(orderquantity)=0
            stop
        endif
    orderline=str(IDNumber)+¬+str(Item)+¬+Description+chr(32)+str(«UnitNumber»)+chr(32)+str(«UnitName»)+¬+
    str(«Sz.»)+¬+"$"+str(«45cost»)+¬+orderquantity+¬+SupplierID+¶

else
    window "45ogscomments.warehouse"
    find Item=extract(itemlist," ",1)
    if info("found")
        getscrap "How many "+Description+" "+str(«UnitNumber»)+ " " +str(«UnitName»)+"?"
        orderquantity=clipboard()
            if val(orderquantity)=0
                stop
            endif
        orderline=str(IDNumber)+¬+str(Item)+¬+Description+¬+str(«UnitNumber»)+" "+str(«UnitName»)+¬+
        str(«Sz.»)+¬+"$" + str(«45cost»)+¬+orderquantity+¬+SupplierID+¶  
    else
    stop
    endif        
endif

//orderquantity=clipboard()
//if val(orderquantity)=0
//    stop
//endif

//orderline=str(IDNumber)+¬+str(Item)+¬+Description+¬+str(«Sz.»)+¬+str(«45cost»)+¬+orderquantity+¬+str(int(«45available»))+¬+SupplierID+¶



window "45ogscomments.linked:Generate-POnew"

forordering=forordering+orderline

;;bigmessage forordering


arrayfilter forordering, intorder, ¶, extract(extract(forordering, ¶, seq()),¬, 1)+¬+extract(extract(forordering, ¶, seq()),¬, 2)+¬+
    extract(extract(forordering, ¶, seq()),¬, 3)+¬+extract(extract(forordering, ¶, seq()),¬, 4)+¬+
    extract(extract(forordering, ¶, seq()),¬, 5)+¬+extract(extract(forordering, ¶, seq()),¬, 6)+¬+
    extract(extract(forordering, ¶, seq()),¬, 7)+¬+¬+¬+¬+extract(extract(forordering, ¶, seq()),¬, 8)
;;bigmessage "intorder"+¶+intorder
clipboard()=intorder

ordersize=extract(extract(intorder,¶, 1),¬,5)

;;bigmessage "ordersize"+¶+ordersize
superobject "fromsupplier", "open"
activesuperobject "Close"

;clipboard()=forordering

___ ENDPROCEDURE .placeorder ___________________________________________________

___ PROCEDURE .superchoicetest _________________________________________________
global vMixCombos


___ ENDPROCEDURE .superchoicetest ______________________________________________

___ PROCEDURE .selectaparent ___________________________________________________
local parento
gettext "Which Item? (XXXX)",parento
select «parent_code»=val(parento)
___ ENDPROCEDURE .selectaparent ________________________________________________

___ PROCEDURE .selectID ________________________________________________________
Item = lookup("45ogscomments.linked","Item",«Item»,"IDNumber",0,0)
Description = lookup("45ogscomments.linked", "Item", «Item», "Description","",0)
___ ENDPROCEDURE .selectID _____________________________________________________

___ PROCEDURE .selectsupplier __________________________________________________
global forordering, itemlist, disorder, seconditemlist, vSupplier
superobject "fromsupplier", "Open"
activesuperobject "Clear"
activesuperobject "Close"
forordering=""
itemlist=""
seconditemlist=""
vSupplier = 0


select SupplierNo_Primary=val(whichsupplier[1,"-"][1,-2])

vSupplier = val(whichsupplier[1,"-"][1,-2])

if info("empty")
else
    arrayselectedbuild itemlist, ¶,"",
    str(Item)[1,-1]+
    rep(chr(32), 31-length(Description[1,30]))+Description[1,30]+
    rep(chr(32), 12-length(«UnitNumber»))+«UnitNumber»+
    rep(chr(32), 18-length(str(«UnitName»)))+str(«UnitName»)+
    rep(chr(32), 10-length(str(«Sz.»)))+str(«Sz.»)+
    rep(chr(32), 8-length(str(«45cost»)))+str(«45cost») +
    rep(chr(32), 34-length(str(«SupplierID»)))+str(«SupplierID»)
endif

 ;; go to comments.warehouse
 
window "45ogscomments.warehouse"

 select SupplierNo_Primary=val(whichsupplier[1,"-"][1,-2])
 if info("empty")
 else
     ;; build a similar array
    arrayselectedbuild seconditemlist, ¶,"",
    str(Item)[1,-1]+
    rep(chr(32), 40-length(Description[1,30]))+Description[1,30]+
    rep(chr(32), 12-length(«UnitNumber»))+«UnitNumber»+
    rep(chr(32), 8-length(str(«UnitName»)))+str(«UnitName»)+
    rep(chr(32), 10-length(str(«Sz.»)))+str(«Sz.»)+
    rep(chr(32), 8-length(str(«45cost»)))+str(«45cost») +
    rep(chr(32), 34-length(str(«SupplierID»)))+str(«SupplierID»)

 endif

 ;; append it to itemlist
 itemlist = itemlist + ¶ + seconditemlist

 window "45ogscomments.linked:Generate-POnew"

superobject "neworder", "FillList"

message "check New Supplier for OG certificate status, special orders, MOQ, etc" 

window "NewSupplier"

select «SupplierIDNumber» = vSupplier 
openform "NewSupplier"




___ ENDPROCEDURE .selectsupplier _______________________________________________

___ PROCEDURE .showtallyorders _________________________________________________
global vItem

vItem = Item

////Select items in tally that contain vItem

Openfile "45ogstally"

Select «Order» contains str(vItem) and sizeof("Status") = 0
___ ENDPROCEDURE .showtallyorders ______________________________________________

___ PROCEDURE .sourcecode ______________________________________________________
local Dictionary, ProcedureList
//this saves your procedures into a variable
 
//step one
saveallprocedures "", Dictionary
clipboard()=Dictionary
//now you can paste those into a text editor and make your changes
STOP
 
//step 2
//this lets you load your changes back in from an editor and put them in
//copy your changed full procedure list back to your clipboard
//now comment out from step one to step 2
//run the procedure one step at a time to load the new list on your clipboard back in
Dictionary=clipboard()
loadallprocedures Dictionary,ProcedureList
message ProcedureList //messages which procedures got changed
___ ENDPROCEDURE .sourcecode ___________________________________________________

___ PROCEDURE .springpuorders __________________________________________________
global vItem

vItem = ""

vItem = Item


////Select items in tally that contain vItem

Openfile "45ogstally"

selectall

Select «Order» contains vItem
SelectWithin «ShipCode» contains "J" or «ShipCode» contains "H" or «ShipCode» contains "T"
SelectWithin «Status» notcontains "Com"

___ ENDPROCEDURE .springpuorders _______________________________________________

___ PROCEDURE .findrepack ______________________________________________________
/***** This macro searches for the repack noted in the "notes" field of CL DB. It extracts
any text after the "#" and searches for that number in the repack DB.  SB 1/3/23 *****/

local vNotes, vRepackNum, vRepackType

vNotes = ""
vRepackNum = ""
vRepackType = ""

vNotes = notes
vRepackType = «Type of Repack»

vRepackNum = after(vNotes,"#")

openfile "45ogsrepack"

select RepackNumber = val(vRepackNum)

case
    vRepackType = "Seed"
    openform "Seed Repack Form"
case
    vRepackType = "Kit"
    openform "Kit Repack Form"
case
    vRepackType contains "Mix"
    openform "Mix Repack Form"
case 
    vRepackType = "Item"
    openform "Item Repack Form"
endcase

    
___ ENDPROCEDURE .findrepack ___________________________________________________

___ PROCEDURE .warehousesupplier _______________________________________________
global forordering, itemlist, disorder, seconditemlist
//superobject "fromsupplier", "Open"
//activesuperobject "Clear"
//activesuperobject "Close"
forordering=""
itemlist=""
seconditemlist=""

window "44ogscomments.warehouse"
select SupplierNo_Primary=val(whichsupplier[1,"¬"])

arrayselectedbuild seconditemlist, ¶,"",Description

 = itemlist + ¶ + seconditemlist

window "44ogscomments.linked:Generate-POnew"

superobject "neworder", "FillList"
___ ENDPROCEDURE .warehousesupplier ____________________________________________

___ PROCEDURE get next ID ______________________________________________________
;; This macro finds the highest permanent ID in the database, and generates the next ID
;; to come after that. The next ID is displayed in a popup message, and is put onto
;; the clipboard, so you can easily paste it.

local all_items, num_items, last_item

arraybuild all_items,¶,"",IDNumber
arraynumericsort all_items,all_items,¶

num_items = arraysize(all_items,¶)
last_item = val(array(all_items,num_items,¶))

clipboard() = last_item + 1
message "Whichever item is the newer one should get an ID of " + str(last_item+1)
___ ENDPROCEDURE get next ID ___________________________________________________

___ PROCEDURE bookpricelines ___________________________________________________
;; this macro generates catalog pricelines for books. It looks like the priceline field
;; is currently being used for something else, so check with Renee before running
;; this as it will wipe out the contents of that field. -SAO 9/7/21

field priceline
formulafill ¬+str(Item)+": "+Description+" ("+str(sz)+"#)/"+pattern(Price, "$#.##")
___ ENDPROCEDURE bookpricelines ________________________________________________

___ PROCEDURE ChkInv ___________________________________________________________
;; This macro is not currently functional, but could in theory be rehabilitated. What it's *trying*
;; to do is loop through all selected records, and if 45available is less than the reorder point,
;; it'll ask you if you want to change the reorder point. There is not currently a field in the
;; database called "reorder point," so it doesn't work.


local stuff
loop
    if «45available»≤«reorder point»
        stuff="Hey! "+str(«Item»)+" "+str(«Sz.»)+"#'s is gettin' low. Would you like to change the reorder threshold?"
        noyes stuff
        if clipboard() contains "y"
            getscrap "New reorder threshold for "+str(«Item»)+"?"
            «reorder point»=val(Clipboard())
        endif
    endif
downrecord
until info("eof")    
___ ENDPROCEDURE ChkInv ________________________________________________________

___ PROCEDURE pricelines _______________________________________________________
;; Fills in catalog pricelines. Don't run this without checking with Renee, as it looks like she is
;; using the priceline field for something else currently. -SAO 9/7/21

field priceline
formulafill ¬+str(Item)+": "+?(«unit note»="",str(«Sz.»)+"# for "+pattern(Price, "$#.##"),«unit note»+" ("+str(«Sz.»)+"#) for "+pattern(Price, "$#.##"))

___ ENDPROCEDURE pricelines ____________________________________________________

___ PROCEDURE tabdown/2 ________________________________________________________
;; a helper macro that triggers equations to recalculate

firstrecord
loop
Cell «»
downrecord
until info("stopped")

___ ENDPROCEDURE tabdown/2 _____________________________________________________

___ PROCEDURE findathing/1 _____________________________________________________
;; This is used in the add-it-up-purchasing form

getscrap "Which Item?"
if val(clipboard())=0
select Description contains clipboard()
field Item
sortup
else
select Item contains clipboard()
field Item
sortup
endif
___ ENDPROCEDURE findathing/1 __________________________________________________

___ PROCEDURE add-it-up-purchasing _____________________________________________
global orderingaddup, soldwta, soldwtb, soldwtc, availwt, reord, soldwtatotal, soldwtbtotal, soldwtctotal, availwttotal, reordtotal

arrayselectedbuild orderingaddup,¶, "45ogscomments.linked",
str(«42sold»)+¬+str(«43sold»)+¬+str(«45sold»)+¬+str(«45available»)+¬+str(«reorder point»)+¬+str(«unit conversion»)+¬+str(«42sold»*«unit conversion»)
+¬+str(«43sold»*«unit conversion»)+¬+str(«45sold»*«unit conversion»)+¬+str(«45available»*«unit conversion»)+¬+str(«reorder point»*«unit conversion»)

;;displaydata orderingaddup

arrayfilter orderingaddup, soldwta,¶,extract(import(),¬,7)
arrayfilter orderingaddup, soldwtb,¶,extract(import(),¬,8)
arrayfilter orderingaddup, soldwtc,¶,extract(import(),¬,9)
arrayfilter orderingaddup, availwt,¶,extract(import(),¬,10)
arrayfilter orderingaddup, reord,¶,extract(import(),¬,11)

;;displaydata reord

arraynumerictotal soldwta,¶, soldwtatotal
arraynumerictotal soldwtb,¶, soldwtbtotal
arraynumerictotal soldwtc,¶, soldwtctotal
arraynumerictotal availwt,¶, availwttotal
arraynumerictotal reord,¶, reordtotal
;;message reordtotal

if info("windowtype") = 15
    drawobjects
endif
___ ENDPROCEDURE add-it-up-purchasing __________________________________________

___ PROCEDURE Check a Pickup/3 _________________________________________________
;; Looks like this macro is supposed to help check the inventory
;; of items needed for a pickup order (or orders).
;; Not sure if this macro is currently being used. -SAO 9/7/21

local nolines
getscrap "How Many Items?"
nolines=val(clipboard())
getscrap "First Item:"
select Item contains clipboard()
if nolines>1
loop
getscrap "Next Item:"
selectadditional Item contains clipboard()
until nolines-1
endif
goform "Inventory for Pickup"
___ ENDPROCEDURE Check a Pickup/3 ______________________________________________

___ PROCEDURE How Much?/4 ______________________________________________________
;; tells you the value of 45available for whatever record you're currently on.
;; Not sure if this macro is being used -SAO 9/7/21

message «45available»
___ ENDPROCEDURE How Much?/4 ___________________________________________________

___ PROCEDURE Generate a Priceline _____________________________________________
;; exports a text file with a priceline for one item (whatever record is active).
;; not sure if this macro is used -SAO 9/7/21

local thisone
thisone=Item
select Item=thisone
export Description+" priceline",Description+"®"+¬+str(Item)+": "+?(«unit note»≠""," "+«unit note»+" ","")+?(Comments="",str(«Sz.»)+"#/"," ("+str(«Sz.»)+"#)/")+pattern(Price, "$#.##")+?(Price≥100,"*","")
selectall
find Item=thisone
___ ENDPROCEDURE Generate a Priceline __________________________________________

___ PROCEDURE pricelines-screwy ________________________________________________
;; not sure whether this macro is used -SAO 9/7/21

;select Listed contains "cat"
field priceline
formulafill ¬+str(«sparetext4»)+": "+?(«unit note»="",str(«Sz.»)+"#/"+pattern(CatalogPrice, "$#.##"),«unit note»+" ("+str(«Sz.»)+"#)/"+pattern(CatalogPrice, "$#.##"))
___ ENDPROCEDURE pricelines-screwy _____________________________________________

___ PROCEDURE delete ___________________________________________________________
;; Loops through the database and deletes all selected records
;; (except for the last one). Seems dangerous! Consider removing? -SAO 9/7/21

yesno "are you sure you want to run the delete macro?"
if clipboard() contains "y"
loop
deleterecord
until info("selected")=1
else
stop
endif
___ ENDPROCEDURE delete ________________________________________________________

___ PROCEDURE checknofa ________________________________________________________
;; I think this macro checks to make sure each successive
;; size code for a given item actually gives you a better deal
;; (by weight). Seems like it's intended to be run while you
;; have all size codes selected for a given seed variety.

local priceone, pricetwo
loop
priceone=CatalogNOFA/ActWt
loop
downrecord
pricetwo=divzero(CatalogNOFA,ActWt)
if pricetwo>.95*priceone
message "check"
stop
endif
priceone=pricetwo
until info("summary")>0
downrecord
until Category notcontains "Seed"
___ ENDPROCEDURE checknofa _____________________________________________________

___ PROCEDURE selectmore _______________________________________________________
global vitemy
loop
vitemy=«parent_code»
selectadditional «parent_code»=vitemy and Listed≠"No"
loop
downrecord
until «parent_code»=vitemy
loop
downrecord 
until «parent_code»≠vitemy
until info("stopped")
___ ENDPROCEDURE selectmore ____________________________________________________

___ PROCEDURE Intrucking change by wt __________________________________________
;; this macro allows you to input the cost of intrucking a 50# bag and automatically
;; calculate the intrucking of the smaller sizes, though you DO need to tab
;; through to get the "base" to update

local tempvariable, costvariable, perpound
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Find «ActWt»=50
gettext "new cost?", costvariable
costvariable=val(costvariable)
«intrucking»=costvariable
perpound=costvariable/50
Field «intrucking»
FormulaFill perpound*«ActWt»
___ ENDPROCEDURE Intrucking change by wt _______________________________________

___ PROCEDURE What's available by the # ________________________________________
local tempvariable, totalvariable

gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «ActWt»*«45available»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries

message "we currently have "+str(totalvariable)+" pounds available"

;This is for automatically calculating how many total pounds we currently have available of a product 8/4/21 RM

___ ENDPROCEDURE What's available by the # _____________________________________

___ PROCEDURE huh ______________________________________________________________
local vNumber, w45CL

vNumber = ""

firstrecord
loop
lookup(file,keyfield,keyvalue,datafield,default,level)ch  ÑGenerate New PO@∂ Ïˇ    ≈  ˇˇ     wasWindow

ˇˇ    wasWindow = info("windowname") 1 ˇˇ: 
   "NewSupplier"m H ˇˇQ    "45ogscomments.warehouse"" l ˇˇs 	   wasWindow
R ~ ˇˇá    "Generate-POnew"ò local wasWindow

wasWindow = info("windowname")

openfile "NewSupplier"
openfile "45ogscomments.warehouse"

window wasWindow

openform "Generate-POnew"
  ÑFY44 Pounds SoldnÏˇ    ≈  ˇˇ     tempvariable, totalvariable
a" ˇˇ*    "What item number? XXXX",lˇˇD    tempvariable) Q ˇˇX    «Item» contains tempvariable u ˇˇ{ 
   "sparemoney3"
0 á ˇˇì    «ActWt»*«44sold»¬ § Â ™ ˇˇµ    totalvariable=sparemoney3Yœ ˇˇ·    {_DatabaseLib}ˇˇ·    {REMOVEALLSUMMARIES}¿‚ ˇˇÍ 9   "in FY44 we've sold "+str(totalvariable)+" pounds so far"Yálocal tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «ActWt»*«44sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "in FY44 we've sold "+str(totalvariable)+" pounds so far"

;This is for automatically calculating how many total pounds we sold FY44 of a product 8/4/21 RM

   ÑFY45 Pounds SoldnÏˇ    ≈  ˇˇ     tempvariable, totalvariable
T" ˇˇ*    "What item number? XXXX", ˇˇD    tempvariable) Q ˇˇX    «Item» contains tempvariable u ˇˇ{ 
   "sparemoney3" 0 á ˇˇì    «ActWt»*«45sold»¬ § Â ™ ˇˇµ    totalvariable=sparemoney3lœ ˇˇ·    {_DatabaseLib}ˇˇ·    {REMOVEALLSUMMARIES}¿‚ ˇˇÍ 9   "in FY45 we've sold "+str(totalvariable)+" pounds so far" álocal tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «ActWt»*«45sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "in FY45 we've sold "+str(totalvariable)+" pounds so far"

;This is for automatically calculating how many total pounds we sold FY45 of a product 8/4/21 RM

 Ï  Ñ	Initial/dˇ Ïˇ    ≈  ˇˇ     finditem, vInitial
Ñ ˇˇ$    "What's the Item# ?"ˇˇ9    finditem=clipboard()) N ˇˇU 
   Item=finditem'Ñd ˇˇo    "What's the initial?"bˇˇÖ    vInitial=val(clipboard())s ü ‚ ˇˇ¢ 
   vInitial=04 ±  ∂  ˇˇø    Initial = vInitial “ ÿ local finditem, vInitial
GetscrapOK "What's the Item# ?"
finditem=clipboard()
Select Item=finditem

GetscrapOK "What's the initial?"
vInitial=val(clipboard())
if vInitial=0
    stop
else
    Initial = vInitial
endif
l  Ñ.huh⁄ Ïˇ    ƒ  ˇˇ     test
ˇˇ 	   test = ""le ˇˇ*    test,(ˇˇ0    ¶,ˇˇ2    "",rˇˇ6 <   rep(chr(32), 20-length(str(«SupplierID»)))+str(«SupplierID»)t ˇˇ    {_DialogAlertLib}hˇˇ    {DISPLAYDATA},ˇˇÄ    test  Ñ Ñ global test
test = ""

arrayselectedbuild test, ¶,"",
rep(chr(32), 20-length(str(«SupplierID»)))+str(«SupplierID»)

displaydata test   Ñ(Purchasing) Ïˇ      ⁄   ÑPO - Open the databases!d Ïˇ       ˇˇ	 
   "NewSupplier"   ˇˇ     "45ogscomments.warehouse"  : ˇˇC    "45OGSpurchasing"cT openfile "NewSupplier"
openfile "45ogscomments.warehouse"
openfile "45OGSpurchasing"v  Ñbringonthebooks/5 Ïˇ    ≈y ˇˇ~     justso, desc, somany

 ˇˇú 	   justso="" ˇˇ¶    desc="" ˇˇÆ    somany=0) ∑ ˇˇæ    Category contains "Book"Ä◊ ˇˇ‡    "Which book came in?" ˇˇˆ    justso=clipboard()+ 	ˇˇ   Description contains justsoAˇˇ2   desc=DescriptionÄCˇˇL   "How many " +desc+ "s came in?" ˇˇl   somany=val(clipboard())  Öˇˇã   "Purchased" ˇˇï   Purchased=Purchased+somany ∞ˇˇ∂
   "On Order"ˇˇ¡   «On Order»=«On Order»-somanyÀﬂˇˇÊ   "Did any other books come in?" ˇˇ   clipboard() contains "Yes"a#*  . 4<;; Looks like this macro is for receiving books. Maybe because book orders
;; don't go through the purchasing database?

local justso, desc, somany

again:
justso=""
desc=""
somany=0
select Category contains "Book"
getscrap "Which book came in?"
justso=clipboard()
selectwithin Description contains justso
desc=Description
getscrap "How many " +desc+ "s came in?"
somany=val(clipboard())

field Purchased
Purchased=Purchased+somany
field «On Order»
«On Order»=«On Order»-somany

Yesno 
"Did any other books come in?"
if clipboard() contains "Yes"
goto again
else 
endif


  ÑAnnual - Pounds NeededdÏˇ    ≈  ˇˇ     tempvariable, totalvariable
h" ˇˇ*    "What item number? XXXX",SˇˇD    tempvariable) Q ˇˇX    «Item» contains tempvariable u ˇˇ{ 
   "sparemoney3"t0 á ˇˇì    «Sz.»*«43sold»¬ ¢ Â ® ˇˇ≥    totalvariable=sparemoney3aÕ ˇˇﬂ    {_DatabaseLib}ˇˇﬂ    {REMOVEALLSUMMARIES}¿‡ ˇˇË 1   "last year we sold "+str(totalvariable)+" pounds" Ålocal tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «Sz.»*«43sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "last year we sold "+str(totalvariable)+" pounds"

;This is for automatically calculating how many total pounds we sold last year of a product 8/4/21 RM dN  ÑAll Seed Sizes and Mixes§ Ïˇ    ≈  ˇˇ     tempvariable
 ˇˇ    "What item number? XXXX", ˇˇ5    tempvariable) B ˇˇI >   «Item» contains tempvariable or «mixkit» contains tempvariableá local tempvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable or «mixkit» contains tempvariable Æ  ÑWhat did we sell/9‘ Ïˇ       f ˇˇ    «add/drop» contains "new"4¿ ˇˇ% &   str(«45sold»)+" sold so far this year" L ‘ ¿Q ˇˇY ^   "43sold: "+str(«43sold»)+¶+"42sold: "+str(«42sold»)+¶+¶+str(«45sold»)+" sold so far this year" ∏ Ω if «add/drop» contains "new"
message str(«45sold»)+" sold so far this year"
else
message "43sold: "+str(«43sold»)+¶+"42sold: "+str(«42sold»)+¶+¶+str(«45sold»)+" sold so far this year"
endifdv  ÑCost Change 50#&smallerpÏˇ    ≈∏ ˇˇΩ &    tempvariable, costvariable, perpound
„ ˇˇÎ    "What item number? XXXX",cˇˇ   tempvariable) ˇˇ   «Item» contains tempvariable* 6ˇˇ;
   «ActWt»=50FˇˇN   "new cost?",ˇˇ[   costvariableˇˇh   costvariable=val(costvariable)ˇˇá   «45cost»=costvariablemˇˇù   perpound=costvariable/50 ∂ˇˇº   "45cost"0 ≈ˇˇ—   perpound*«ActWt»„;; this macro allows you to input the cost of a 50# bag and automatically
;; calculate the cost of the smaller sizes, though you DO need to tab
;; through to get the "base" to update

local tempvariable, costvariable, perpound
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Find «ActWt»=50
gettext "new cost?", costvariable
costvariable=val(costvariable)
«45cost»=costvariable
perpound=costvariable/50
Field «45cost»
FormulaFill perpound*«ActWt»

 t  ÑCost Change 25#&smallerpÏˇ    ≈π ˇˇæ &    tempvariable, costvariable, perpound
‰ ˇˇÏ    "What item number? XXXX", ˇˇ   tempvariable) ˇˇ   «Item» contains tempvariable* 7ˇˇ<
   «ActWt»=25GˇˇO   "new cost?",ˇˇ\   costvariableˇˇi   costvariable=val(costvariable)ˇˇà   «45cost»=costvariable ˇˇû   perpound=costvariable/25 ∑ˇˇΩ   "45cost"0 ∆ˇˇ“   perpound*«ActWt»‚;; this macro allows you to input the cost of a 25# bag and automatically
;; calculate the cost of the smaller sizes, though you DO need to tab
;; through to get the "base" to update


local tempvariable, costvariable, perpound
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Find «ActWt»=25
gettext "new cost?", costvariable
costvariable=val(costvariable)
«45cost»=costvariable
perpound=costvariable/25
Field «45cost»
FormulaFill perpound*«ActWt»  ÑFY43 sold - total feetnÏˇ    ≈  ˇˇ     tempvariable, totalvariable
f" ˇˇ*    "What item number? XXXX",lˇˇD    tempvariable) Q ˇˇX    «Item» contains tempvariable u ˇˇ{ 
   "sparemoney3"ˇ0 á ˇˇì    «unit conversion»*«43sold»¬ Æ Â ¥ ˇˇø    totalvariable=sparemoney3eŸ ˇˇÎ    {_DatabaseLib}ˇˇÎ    {REMOVEALLSUMMARIES}¿Ï ˇˇÙ /   "last year we sold "+str(totalvariable)+" feet"Nálocal tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «unit conversion»*«43sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "last year we sold "+str(totalvariable)+" feet"

;This is for automatically calculating how many total feet we sold last year of a product 10/21 RM""  ÑAvailable - by the foottÏˇ    ≈  ˇˇ     tempvariable, totalvariable
u" ˇˇ*    "What item number? XXXX",bˇˇD    tempvariable) Q ˇˇX    «Item» contains tempvariable u ˇˇ{ 
   "sparemoney3" 0 á ˇˇì    «unit conversion»*«45available»"¬ ≥ Â π ˇˇƒ    totalvariable=sparemoney3lﬁ ˇˇ    {_DatabaseLib}ˇˇ    {REMOVEALLSUMMARIES}¿Ò ˇˇ˘ /   "we have "+str(totalvariable)+" feet available"målocal tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «unit conversion»*«45available»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "we have "+str(totalvariable)+" feet available"

;This is for automatically calculating how many total feet we have available of a product 10/21 RM  ÑGenerate POdÑ Ïˇ       ˇˇ	 
   "NewSupplier"   ˇˇ     "45ogscomments.warehouse"r : ˇˇA    "45ogscomments.linked"R X ˇˇa    "Generate-POnew"q openfile "NewSupplier"
openfile "45ogscomments.warehouse"
window "45ogscomments.linked"
openform "Generate-POnew"iZ   ÑPricing Inquiryo$ Ïˇ    R   ˇˇ	    "Pricing Inquiry" openform "Pricing Inquiry"
 ¯  ÑFY42 - Pounds SoldbÏˇ    ≈  ˇˇ     tempvariable, totalvariable
a" ˇˇ*    "What item number? XXXX",lˇˇD    tempvariable) Q ˇˇX    «Item» contains tempvariable u ˇˇ{ 
   "sparemoney3"a0 á ˇˇì    «Sz.»*«42sold»¬ ¢ Â ® ˇˇ≥    totalvariable=sparemoney3nÕ ˇˇﬂ    {_DatabaseLib}ˇˇﬂ    {REMOVEALLSUMMARIES}¿‡ ˇˇË /   "in FY42 we sold "+str(totalvariable)+" pounds"lylocal tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «Sz.»*«42sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "in FY42 we sold "+str(totalvariable)+" pounds"

;This is for automatically calculating how many total pounds we sold FY42 of a product 8/4/21 RMm  ÑAll Supplier Parent CodesÏÏˇ    ≈  ˇˇ     vSupplier, vParentcode

pÑÏ ˇˇ˜ #   "What is the name of the supplier?"Lˇˇ   vSupplier=str(clipboard())) 6ˇˇ=b   «Supplier» contains vSupplier

///Expand the search to include all size codes within the selection †ˇˇ¶
   "parent_code"r¢ ¥‰ ªˇˇ»   vParentcode=«parent_code»1, ‚ˇˇÛ,   str(«parent_code») contains str(vParentcode)ë! ) 2¢ˇˇ>   «parent_code»=vParentcode4„ ` t∆ˇˇÇ   vParentcode=""ˇˇô   vParentcode=«parent_code» , ªˇˇÃ´   str(«parent_code») contains str(vParentcode)
        
///Below piece of code put in because after running the "else" statement the first record
///Would be selected again."* ÅˇˇÜ    val(vParentcode) = «parent_code»„ Ø ¡ «pˇˇÕ   info("stopped")píﬁ‰ Èılocal vSupplier, vParentcode

///vSupplier will be defined as the name of the supplier
///vParentcode will be defined as the parent code of items
///Search within the Supplier field that contains what is entered in the GetScrap prompt

GetscrapOK "What is the name of the supplier?"
vSupplier=str(clipboard())
Select «Supplier» contains vSupplier

///Expand the search to include all size codes within the selection
Field «parent_code»
Sortup
Firstrecord

vParentcode=«parent_code»
SelectAdditional str(«parent_code») contains str(vParentcode)

noshow

loop
    if 
        «parent_code»=vParentcode
        downrecord     
    else 
        vParentcode=""
        vParentcode=«parent_code»
        SelectAdditional str(«parent_code») contains str(vParentcode)
        
///Below piece of code put in because after running the "else" statement the first record
///Would be selected again.

        find val(vParentcode) = «parent_code»
        downrecord   
    endif
until info("stopped")

endnoshow

firstrecord
_   Ñ(Inventory)@ Ïˇ      å  ÑUpdate Tally Numbers\Ïˇ      ˇˇ 
   {_UtilityLib}dˇˇ    {REMEMBERWINDOW}  ˇˇ    "45orderedogs"( ˇˇ6 
   {_UtilityLib}(ˇˇ6    {ORIGINALWINDOW} 8 ˇˇ>    ("45soldtally")e) N ˇˇU E   «» <> lookupselected("45orderedogs","IDNumber", IDNumber, "qty", 0,1)0+ ú ˇˇ© J   IDNumber=lookupselected("45orderedogs","IDNumber",IDNumber,"IDNumber",0,1) ı ˇˇ˚    ("45soldtally");0 ˇˇ?   lookupselected("45orderedogs","IDNumber", IDNumber, "qty", 0,1)o Üˇˇå   ("45tallyfilled")N) ûˇˇ•C   «» <> lookupselected("45orderedogs","IDNumber",IDNumber,"fill",0,1),+ Íˇˇ˜J   IDNumber=lookupselected("45orderedogs","IDNumber",IDNumber,"IDNumber",0,1) CˇˇI   ("45tallyfilled")N0 [ˇˇg=   lookupselected("45orderedogs","IDNumber",IDNumber,"fill",0,1)n ≥ˇˇ∫   "45orderedogs" …÷ˇˇ‰
   {_UtilityLib}Lˇˇ‰   {ORIGINALWINDOW}¿ÊˇˇÓ#   "Finished Tally Sold/Filled Update"arememberwindow
openfile "45orderedogs"

originalwindow

field ("45soldtally")
select «» <> lookupselected("45orderedogs","IDNumber", IDNumber, "qty", 0,1)

Selectwithin IDNumber=lookupselected("45orderedogs","IDNumber",IDNumber,"IDNumber",0,1)

field ("45soldtally")
formulafill lookupselected("45orderedogs","IDNumber", IDNumber, "qty", 0,1)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Field ("45tallyfilled")
Select «» <> lookupselected("45orderedogs","IDNumber",IDNumber,"fill",0,1)

selectwithin IDNumber=lookupselected("45orderedogs","IDNumber",IDNumber,"IDNumber",0,1)

Field ("45tallyfilled")
formulafill lookupselected("45orderedogs","IDNumber",IDNumber,"fill",0,1)

;;SelectAll

window "45orderedogs"
closewindow

originalwindow

message "Finished Tally Sold/Filled Update"e\  Ñrebuild_repack_numbersXÏˇ      ˇˇ 
   {_UtilityLib}cˇˇ    {REMEMBERWINDOW}  ˇˇ    "repack_vertical_file"0 ˇˇ> 
   {_UtilityLib}eˇˇ>    {ORIGINALWINDOW} @ ˇˇF    ("45repack")) S ˇˇZ J   «» <> lookupselected("repack_vertical_file","IDNumber",IDNumber,"Net",0,0)+ ¶ ˇˇ≥ R   IDNumber=lookupselected("repack_vertical_file","IDNumber",IDNumber,"IDNumber",0,0) ˇˇ
   ("45repack")0 ˇˇ&D   lookupselected("repack_vertical_file","IDNumber",IDNumber,"Net",0,0)∆lˇˇq   "tabdown/2"e  ~ âˇˇê   "repack_vertical_file" ß¥ˇˇ¬
   {_UtilityLib}lˇˇ¬   {ORIGINALWINDOW}¿ƒˇˇÃ   "Finished Repack Update"‰rememberwindow
openfile "repack_vertical_file"

originalwindow

Field ("45repack")
Select «» <> lookupselected("repack_vertical_file","IDNumber",IDNumber,"Net",0,0)

selectwithin IDNumber=lookupselected("repack_vertical_file","IDNumber",IDNumber,"IDNumber",0,0)

Field ("45repack")
formulafill lookupselected("repack_vertical_file","IDNumber",IDNumber,"Net",0,0)

call "tabdown/2"

;SelectAll
window "repack_vertical_file"
closewindow

originalwindow

message "Finished Repack Update"Z	  Ñwalkin_sales_update ¬Ïˇ      ˇˇ 
   {_UtilityLib}-ˇˇ    {REMEMBERWINDOW} S ˇˇ\    "45transfers_vertical_ogs"x ˇˇÜ 
   {_UtilityLib}"ˇˇÜ    {ORIGINALWINDOW} á ˇˇç    ("45transfers") ) ù ˇˇ§ N   «» <> lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"qty",0,1)+ Ù ˇˇV   IDNumber=lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"IDNumber",0,1) Yˇˇ_   ("45transfers")s0 oˇˇ{H   lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"qty",0,1) ≈ˇˇÃ   "45transfers_vertical_ogs" ÁÙˇˇ
   {_UtilityLib} ˇˇ   {ORIGINALWINDOW} ˇˇ

   "45transfers" ∆ˇˇ   "tabdown/2"   *é * 5ˇˇ>   "45walkin_vertical_ogs"tùˇˇ´
   {_UtilityLib}tˇˇ´   {ORIGINALWINDOW} ¨ˇˇ≤   ("45soldwalkin")) √ˇˇ K   «» <> lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"qty",0,1)"+ ˇˇ$S   IDNumber=lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"IDNumber",0,1)m yˇˇ   ("45soldwalkin")0 êˇˇúE   lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"qty",0,1)" Ôˇˇˆ   "45walkin_vertical_ogs"t ˇˇ)
   {_UtilityLib}tˇˇ)   {ORIGINALWINDOW} +ˇˇ1   "45soldwalkin"∆@ˇˇE   "tabdown/2"t  R¿RˇˇZ   "Finished Walk-in Sales Update"ezrememberwindow

;;--------------------------- transfers -----------------------;;

openfile "45transfers_vertical_ogs"

originalwindow
Field ("45transfers")
Select «» <> lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

selectwithin IDNumber=lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"IDNumber",0,1)

Field ("45transfers")
formulafill lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

window "45transfers_vertical_ogs"
closewindow

originalwindow

field «45transfers»
call "tabdown/2"

SelectAll

openfile "45walkin_vertical_ogs"

;;--------------------------- normal sales -----------------------;;

originalwindow
Field ("45soldwalkin")
Select «» <> lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

selectwithin IDNumber=lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"IDNumber",0,1)

Field ("45soldwalkin")
formulafill lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

;;SelectAll
window "45walkin_vertical_ogs"
closewindow

originalwindow

field «45soldwalkin»
call "tabdown/2"

message "Finished Walk-in Sales Update"
ÿ  Ñ	add it up@ 	Ïˇ    é   ë ) D ˇˇK F   «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds») í Ü ˇˇï 
   info("empty")ré ß  ± ¬  ∫ ˇˇ¿    ("45soldtally")y∆‘ ˇˇŸ    "tabdown/2"   Â  Â ) Ï ˇˇÛ F   «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds») ::ˇˇ=
   info("empty")4é O Yv bˇˇh   ("45soldwalkin")∆}ˇˇÇ   "tabdown/2"4  é é) ïˇˇúF   «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds») „ÓˇˇÊ
   info("empty"))é ¯ * ˇˇ   ("45transfers") ∆%ˇˇ*   "tabdown/2")  6 6) çˇˇît   «45available» <> (Initial+Purchased+«45repack»-«45soldtally»-«45soldwalkin»-«45soldseeds»-«45transfers»+Adjustments) 	–ˇˇ
   info("empty")«é  ( 1ˇˇ7   ("45repack")∆HˇˇM   "tabdown/2"d  Y Y) ™ˇˇ±>   «45inhouse» <> («45available»+(«45soldtally»-«45tallyfilled»)) xˇˇÛ
   info("empty")-é  ¥ ˇˇ   ("45soldtally")-∆2ˇˇ7   "tabdown/2"
  C C) JˇˇQ>   «45inhouse» <> («45available»+(«45soldtally»-«45tallyfilled»)) ê$ˇˇì
   info("empty") é • Øb ∏ˇˇæ   ("45tallyfilled")/∆‘ˇˇŸ   "tabdown/2"o  Â Âé 5 ?ˇˇH   "45ogscomments"4 Xˇˇa   "&&45ogscomments.linked") {ˇˇÇ   «can be added up» contains "Y" °ˇˇß
   "parent_code"d¶ µ æˇˇƒ   "TotalPoundsToSell"-0 ÷ˇˇ‚   «unit conversion»*«45available»l¬ ⁄ 	ˇˇ   "1"o ˇˇ#   "TotalUnitsToSell"0 6ˇˇB   TotalPoundsToSellb⁄ Uˇˇe   "Data" mˇˇs   "TotalPoundsToSell"p∞ áˇˇå   ""⁄ èˇˇü   "1"o0 §ˇˇ∞   TotalUnitsToSell⁄ ¬ˇˇ“   "Data" Ÿˇˇﬂ   "TotalPoundsToSell"o∂ Û®  ˇˇ   "7"m ˇˇ   "TotalUnitsToSell"0 .ˇˇ:<   TotalPoundsToSell/?(«unit conversion»>0,«unit conversion»,1)) xˇˇ   «can be added up» contains "N" ûˇˇ§   "TotalUnitsToSell"0 ∑ˇˇ√
   «45available»eí“ ›ˇˇ‰   "45ogscomments.linked" ¸ˇˇ   "TotalUnitsToSell") ˇˇP   «» <> lookupselected("45ogscomments","IDNumber",IDNumber,"TotalUnitsToSell",0,0)0 mˇˇyB   lookup("45ogscomments","IDNumber",IDNumber,"TotalUnitsToSell",0,0) Ωˇˇ√   "TotalPoundsToSell"S) ◊ˇˇﬁQ   «» <> lookupselected("45ogscomments","IDNumber",IDNumber,"TotalPoundsToSell",0,0)F0 0	ˇˇ<	C   lookup("45ogscomments","IDNumber",IDNumber,"TotalPoundsToSell",0,0)Fé Å	 å	ˇˇí	   "Item"¢ ô	¿°	ˇˇ©	   "Finished with Add-it-up!"√	Selectall

NoShow

;;-------------------SOLD----------------------

Select «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds»)
if info("empty")
    SelectAll
else
    Field ("45soldtally")
    call "tabdown/2"
endif

Select «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds»)
if info("empty")
    SelectAll
else
    Field ("45soldwalkin")
    call "tabdown/2"
endif

Select «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds»)
if info("empty")
    SelectAll
else
    Field ("45transfers")
    call "tabdown/2"
endif

;; ---------------------------------AVAILABLE---------------------------------

Select «45available» <> (Initial+Purchased+«45repack»-«45soldtally»-«45soldwalkin»-«45soldseeds»-«45transfers»+Adjustments)
if info("empty")
    SelectAll
else
    Field ("45repack")
    call "tabdown/2"
endif

;; --------------------------------IN HOUSE-----------------------------

Select «45inhouse» <> («45available»+(«45soldtally»-«45tallyfilled»))
if info("empty")
    SelectAll
else
    Field ("45soldtally")
    call "tabdown/2"
endif

Select «45inhouse» <> («45available»+(«45soldtally»-«45tallyfilled»))
if info("empty")
    SelectAll
else
    Field ("45tallyfilled")
    call "tabdown/2"
endif

;; -------------------------------- add it up -------------------------

SelectAll
openfile "45ogscomments"
openfile "&&45ogscomments.linked"

Select «can be added up» contains "Y"
Field "parent_code"
GroupUp

Field TotalPoundsToSell
FormulaFill «unit conversion»*«45available»
Total

CollapseToLevel "1"
Field "TotalUnitsToSell"
FormulaFill TotalPoundsToSell

CollapseToLevel "Data"

Field "TotalPoundsToSell"
Fill ""
CollapseToLevel "1"

FormulaFill TotalUnitsToSell

CollapseToLevel "Data"
Field "TotalPoundsToSell"
PropagateUP

RemoveSummaries "7"

Field "TotalUnitsToSell"
FormulaFill TotalPoundsToSell/?(«unit conversion»>0,«unit conversion»,1)

Select «can be added up» contains "N"
Field "TotalUnitsToSell"
FormulaFill «45available»

EndNoShow

window "45ogscomments.linked"

Field "TotalUnitsToSell"
Select «» <> lookupselected("45ogscomments","IDNumber",IDNumber,"TotalUnitsToSell",0,0)
FormulaFill lookup("45ogscomments","IDNumber",IDNumber,"TotalUnitsToSell",0,0)

Field "TotalPoundsToSell"
Select «» <> lookupselected("45ogscomments","IDNumber",IDNumber,"TotalPoundsToSell",0,0)
FormulaFill lookup("45ogscomments","IDNumber",IDNumber,"TotalPoundsToSell",0,0)

SelectAll

Field "Item"
Sortup

message "Finished with Add-it-up!"_h  ÑNegative Report@Ç Ïˇ     N ˇˇT    "Item"¢ ]  i ˇˇo    "Item") v ˇˇ} $   «45available» < 0 or «45inhouse» < 0R £ ˇˇ¨    "Negative Report""Kø / Ã ;; searches for items that are less than 0 for Stasha to go and investigate.

Field Item
    sortup
    
Field "Item" Select «45available» < 0 or «45inhouse» < 0

openform "Negative Report"

Print Dialog
  ÑHide Most FieldsîÏˇ    ≈  ˇˇ     fieldstoshow
ˇˇ $  fieldstoshow = "Listed"+¶+"IDNumber"+¶+"Item"+¶+"Description"+¶+"Sz."+¶+"ActWt"+¶+"Comments"+¶+"notes"+¶+"45available"+¶+"TotalPoundsToSell"+¶+"TotalUnitsToSell"+¶+"Initial"+¶+"Adjustments"+¶+"45sold"+¶+"45inhouse"+¶+"45tallyfilled"+¶+"45soldtally"+¶+"45soldwalkin"+¶+"45repack"+¶+"unit note"9ˇˇH   {_DatabaseLib}ˇˇH   {SHOWTHESEFIELDS},ˇˇI   fieldstoshowUlocal fieldstoshow
fieldstoshow = "Listed"+¶+"IDNumber"+¶+"Item"+¶+"Description"+¶+"Sz."+¶+"ActWt"+¶+"Comments"+¶+"notes"+¶+"45available"+¶+"TotalPoundsToSell"+¶+"TotalUnitsToSell"+¶+"Initial"+¶+"Adjustments"+¶+"45sold"+¶+"45inhouse"+¶+"45tallyfilled"+¶+"45soldtally"+¶+"45soldwalkin"+¶+"45repack"+¶+"unit note"

showthesefields fieldstoshowen   Ñ	Inquiry/I 2 Ïˇ       ˇˇ    "Item"¢  R  ˇˇ 	   "Inquiry"e' field Item
sortup
openform "Inquiry"


d  ÑZero Report@fÏˇ     t ˇˇz    "Item"¢ Å ) â ˇˇê ñ   («45available» = 0 or «45inhouse» = 0) and «Comments» notcontains "clearance" and «notes» notcontains "clearance" and «Comments» notcontains "service"À(ˇˇ.4   "Only show items with no inventory and no comments?" gFˇˇs   clipboard() contains "yes"+ öˇˇß   «Comments» = ""ˇ ªR ¬ˇˇÀ
   "Zero Report"oK⁄/ Á;;Searches for items that have a 0 "45Available" to make sure that comments are filled in and counts are accurate.

Field "Item"
sortup

Select («45available» = 0 or «45inhouse» = 0) and «Comments» notcontains "clearance" and «notes» notcontains "clearance" and «Comments» notcontains "service"

yesno "Only show items with no inventory and no comments?"
    if 
        clipboard() contains "yes"
            selectwithin «Comments» = ""
    endif

openform "Zero Report"

print dialog
0Ï   Ñ
Update Repack@p Ïˇ       ˇˇ 
   "45repack")  ˇˇ    «45repack» <> 0 ‰ )  5 Ö: ˇˇ?    «»„ B  M @ ˇˇS    info("stopped") c Field "45repack"
Select «45repack» <> 0

firstrecord
loop
Cell «»
downrecord
until info("stopped")
   Ñ	Priceline@å Ïˇ       ˇˇ 
   "add/drop")  ˇˇ c   «priceline» notcontains |||21||| or sizeof("priceline") = 0 and («add/drop» notcontains |||drop|||) | Field "add/drop"
Select «priceline» notcontains |||21||| or sizeof("priceline") = 0 and («add/drop» notcontains |||drop|||) t   Ñdrop4 Ïˇ       ˇˇ 
   "IDNumber"¢  A ˇˇ)    IDNumber1 Field «IDNumber»
Sortup
Selectduplicates IDNumber ™  ÑSale History‹ Ïˇ    ≈ ˇˇ     fieldstoshow
ˇˇ, l   fieldstoshow = "Item"+¶+"Description"+¶+"45available"+¶+"45sold"+¶+"44sold"+¶+"43sold"+¶+"42sold"+¶+"41sold"ö ˇˇ©    {_DatabaseLib}ˇˇ©    {SHOWTHESEFIELDS},ˇˇ™    fieldstoshow∑ ;;Searches sale history

Local fieldstoshow
fieldstoshow = "Item"+¶+"Description"+¶+"45available"+¶+"45sold"+¶+"44sold"+¶+"43sold"+¶+"42sold"+¶+"41sold"

showthesefields fieldstoshow
 J  ÑArchive Item Search ÄÏˇ     D ˇˇJ    "Item") Q ˇˇX a   «Listed» = |||No||| and «add/drop» = |||historical||| and «45available» <= 0 and «45inhouse» <= 0  º ÷ ˇˇø 
   info("empty") ¿— ˇˇŸ #   "No items that need to be archived"  ˝ + ˇˇ   «45soldtally» > «45tallyfilled» 1Fˇˇ4
   info("empty") ¿FˇˇN   "No outstanding orders"  g¿pˇˇx-   "Next Step- Run Archive Selected Items macro" ¨;Search for items that may be able to be moved to comments.archive

Field "Item"
Select «Listed» = |||No||| and «add/drop» = |||historical||| and «45available» <= 0 and «45inhouse» <= 0 

if info("empty")
    message "No items that need to be archived"
endif
Selectwithin «45soldtally» > «45tallyfilled»

if info("empty")
    message "No outstanding orders"
 endif
 
 message "Next Step- Run Archive Selected Items macro"






T  ÑArchive Selected Items|Ïˇ    …  ˇˇ	 R   "Careful! Make sure you have only the items you want appended selected. Continue?" ` ê ˇˇc    clipboard() = "Cancel"4 Ç  ã  í ˇˇõ    "Comments.archive" Ø ˇˇ∏    "++45ogscomments.linked" ” ˇˇ⁄    "45ogscomments.linked"Ã!ˇˇ''   "Do you wish to delete appended items?"s R|ˇˇU   clipboard() = "Yes"i lÂ tv Ç íNˇˇò   info("selected") = 1 ≠∏okcancel "Careful! Make sure you have only the items you want appended selected. Continue?"
    if clipboard() = "Cancel"
        stop
    endif
 openfile "Comments.archive"
 openfile "++45ogscomments.linked" 
 window "45ogscomments.linked"

;Delete appended records from Comments.Linked

noyes "Do you wish to delete appended items?"  

if clipboard() = "Yes"
   Loop
   Lastrecord
   Deleterecord
   until info("selected") = 1
endif





z   ÑSearch< Ïˇ       ˇˇ    "Item") 
 ˇˇ    «45available» > «43sold»- Field "Item"
Select «45available» > «43sold»  N   ÑItem History  Ïˇ    R   ˇˇ	    "Item History" Openform "Item History"   ÑZero Audit Form@Ñ Ïˇ    R   ˇˇ	    "Blank Zero Audit Form" ¿! ˇˇ) G   "When printing make sure to select 1 from 1 for print options, NOT all" Kq / ~ Openform "Blank Zero Audit Form"
M
___ ENDPROCEDURE huh ___________________________________________________________

___ PROCEDURE Generate New PO __________________________________________________
global wCL, wNewSupplier, wCommentsWarehouse

wCL = ""
wNewSupplier = ""
wCommentsWarehouse = ""

wCL = info("windowname")

openfile "NewSupplier"
    
    wNewSupplier = info("windowname")
    
openfile "45ogscomments.warehouse"

    wCommentsWarehouse = info("windowname")
    
window wCL

openform "Generate-POnew"

___ ENDPROCEDURE Generate New PO _______________________________________________

___ PROCEDURE FY44 Pounds Sold _________________________________________________
local tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «ActWt»*«44sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "in FY44 we've sold "+str(totalvariable)+" pounds so far"

;This is for automatically calculating how many total pounds we sold FY44 of a product 8/4/21 RM


___ ENDPROCEDURE FY44 Pounds Sold ______________________________________________

___ PROCEDURE FY45 Pounds Sold _________________________________________________
local tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «ActWt»*«45sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "in FY45 we've sold "+str(totalvariable)+" pounds so far"

;This is for automatically calculating how many total pounds we sold FY45 of a product 8/4/21 RM


___ ENDPROCEDURE FY45 Pounds Sold ______________________________________________

___ PROCEDURE Initial/d ________________________________________________________
local finditem, vInitial
GetscrapOK "What's the Item# ?"
finditem=clipboard()
Select Item=finditem

GetscrapOK "What's the initial?"
vInitial=val(clipboard())
if vInitial=0
    stop
else
    Initial = vInitial
endif

___ ENDPROCEDURE Initial/d _____________________________________________________

___ PROCEDURE Test _____________________________________________________________
sendoneemail "slbaldwin91@gmail.com", "stasha@mainebeefalliance.net", "Test", "This is a test"
___ ENDPROCEDURE Test __________________________________________________________

___ PROCEDURE Export Pan Macros ________________________________________________
local Dictionary, ProcedureList
//this saves your procedures into a variable
 
 
//step one
saveallprocedures "", Dictionary
clipboard()=Dictionary
//now you can paste those into a text editor and make your changes

___ ENDPROCEDURE Export Pan Macros _____________________________________________

___ PROCEDURE Import Pan Macros ________________________________________________
local Dictionary, ProcedureList

//step 2
//this lets you load your changes back in from an editor and put them in
//copy your changed full procedure list back to your clipboard
//now comment out from step one to step 2
//run the procedure one step at a time to load the new list on your clipboard back in
Dictionary=clipboard()
loadallprocedures Dictionary,ProcedureList
message ProcedureList //messages which procedures got changed
___ ENDPROCEDURE Import Pan Macros _____________________________________________

___ PROCEDURE (Purchasing) _____________________________________________________

___ ENDPROCEDURE (Purchasing) __________________________________________________

___ PROCEDURE All Seed Sizes and Mixes _________________________________________
local tempvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable or «mixkit» contains tempvariable
___ ENDPROCEDURE All Seed Sizes and Mixes ______________________________________

___ PROCEDURE All Supplier Parent Codes ________________________________________
local vSupplier, vParentcode

///vSupplier will be defined as the name of the supplier
///vParentcode will be defined as the parent code of items
///Search within the Supplier field that contains what is entered in the GetScrap prompt

GetscrapOK "What is the name of the supplier?"
vSupplier=str(clipboard())
Select «Supplier» contains vSupplier

///Expand the search to include all size codes within the selection
Field «parent_code»
Sortup
Firstrecord

vParentcode=«parent_code»
SelectAdditional str(«parent_code») contains str(vParentcode)

noshow

loop
    if 
        «parent_code»=vParentcode
        downrecord     
    else 
        vParentcode=""
        vParentcode=«parent_code»
        SelectAdditional str(«parent_code») contains str(vParentcode)
        
///Below piece of code put in because after running the "else" statement the first record
///Would be selected again.

        find val(vParentcode) = «parent_code»
        downrecord   
    endif
until info("stopped")

endnoshow

firstrecord

___ ENDPROCEDURE All Supplier Parent Codes _____________________________________

___ PROCEDURE Annual - Pounds Needed ___________________________________________
local tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «Sz.»*«43sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "last year we sold "+str(totalvariable)+" pounds"

;This is for automatically calculating how many total pounds we sold last year of a product 8/4/21 RM 
___ ENDPROCEDURE Annual - Pounds Needed ________________________________________

___ PROCEDURE Available - by the foot __________________________________________
local tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «unit conversion»*«45available»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "we have "+str(totalvariable)+" feet available"

;This is for automatically calculating how many total feet we have available of a product 10/21 RM
___ ENDPROCEDURE Available - by the foot _______________________________________

___ PROCEDURE bringonthebooks/5 ________________________________________________
;; Looks like this macro is for receiving books. Maybe because book orders
;; don't go through the purchasing database?

local justso, desc, somany

again:
justso=""
desc=""
somany=0
select Category contains "Book"
getscrap "Which book came in?"
justso=clipboard()
selectwithin Description contains justso
desc=Description
getscrap "How many " +desc+ "s came in?"
somany=val(clipboard())

field Purchased
Purchased=Purchased+somany
field «On Order»
«On Order»=«On Order»-somany

Yesno 
"Did any other books come in?"
if clipboard() contains "Yes"
goto again
else 
endif



___ ENDPROCEDURE bringonthebooks/5 _____________________________________________

___ PROCEDURE Cost Change 25#&smaller __________________________________________
;; this macro allows you to input the cost of a 25# bag and automatically
;; calculate the cost of the smaller sizes, though you DO need to tab
;; through to get the "base" to update


local tempvariable, costvariable, perpound
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Find «ActWt»=25
gettext "new cost?", costvariable
costvariable=val(costvariable)
«45cost»=costvariable
perpound=costvariable/25
Field «45cost»
FormulaFill perpound*«ActWt»
___ ENDPROCEDURE Cost Change 25#&smaller _______________________________________

___ PROCEDURE Cost Change 50#&smaller __________________________________________
;; this macro allows you to input the cost of a 50# bag and automatically
;; calculate the cost of the smaller sizes, though you DO need to tab
;; through to get the "base" to update

local tempvariable, costvariable, perpound
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Find «ActWt»=50
gettext "new cost?", costvariable
costvariable=val(costvariable)
«45cost»=costvariable
perpound=costvariable/50
Field «45cost»
FormulaFill perpound*«ActWt»


___ ENDPROCEDURE Cost Change 50#&smaller _______________________________________

___ PROCEDURE FY42 - Pounds Sold _______________________________________________
local tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «Sz.»*«42sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "in FY42 we sold "+str(totalvariable)+" pounds"

;This is for automatically calculating how many total pounds we sold FY42 of a product 8/4/21 RM
___ ENDPROCEDURE FY42 - Pounds Sold ____________________________________________

___ PROCEDURE FY43 sold - total feet ___________________________________________
local tempvariable, totalvariable
gettext "What item number? XXXX", tempvariable
Select «Item» contains tempvariable
Field sparemoney3
FormulaFill «unit conversion»*«43sold»
Total
lastrecord
totalvariable=sparemoney3
removeallsummaries
message "last year we sold "+str(totalvariable)+" feet"

;This is for automatically calculating how many total feet we sold last year of a product 10/21 RM
___ ENDPROCEDURE FY43 sold - total feet ________________________________________

___ PROCEDURE Generate PO ______________________________________________________
openfile "NewSupplier"
openfile "45ogscomments.warehouse"
window "45ogscomments.linked"
openform "Generate-POnew"
___ ENDPROCEDURE Generate PO ___________________________________________________

___ PROCEDURE PO - Open the databases! _________________________________________
openfile "NewSupplier"
openfile "45ogscomments.warehouse"
openfile "45OGSpurchasing"
___ ENDPROCEDURE PO - Open the databases! ______________________________________

___ PROCEDURE Pricing Inquiry __________________________________________________
openform "Pricing Inquiry"

___ ENDPROCEDURE Pricing Inquiry _______________________________________________

___ PROCEDURE What did we sell/9 _______________________________________________
if «add/drop» contains "new"
message str(«45sold»)+" sold so far this year"
else
message "43sold: "+str(«43sold»)+¶+"42sold: "+str(«42sold»)+¶+¶+str(«45sold»)+" sold so far this year"
endif
___ ENDPROCEDURE What did we sell/9 ____________________________________________

___ PROCEDURE (Inventory) ______________________________________________________

___ ENDPROCEDURE (Inventory) ___________________________________________________

___ PROCEDURE Add It Up ________________________________________________________
Selectall

NoShow

;;-------------------SOLD----------------------

Select «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds»)
if info("empty")
    SelectAll
else
    Field ("45soldtally")
    call "tabdown/2"
endif

Select «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds»)
if info("empty")
    SelectAll
else
    Field ("45soldwalkin")
    call "tabdown/2"
endif

Select «45sold» <> («45soldtally»+«45soldwalkin»+«45transfers»+«45soldseeds»)
if info("empty")
    SelectAll
else
    Field ("45transfers")
    call "tabdown/2"
endif

;; ---------------------------------AVAILABLE---------------------------------

Select «45available» <> (Initial+Purchased+«45repack»-«45soldtally»-«45soldwalkin»-«45soldseeds»-«45transfers»+Adjustments)
if info("empty")
    SelectAll
else
    Field ("45repack")
    call "tabdown/2"
endif

;; --------------------------------IN HOUSE-----------------------------

Select «45inhouse» <> («45available»+(«45soldtally»-«45tallyfilled»))
if info("empty")
    SelectAll
else
    Field ("45soldtally")
    call "tabdown/2"
endif

Select «45inhouse» <> («45available»+(«45soldtally»-«45tallyfilled»))
if info("empty")
    SelectAll
else
    Field ("45tallyfilled")
    call "tabdown/2"
endif

;; -------------------------------- add it up -------------------------

SelectAll
openfile "45ogscomments"
openfile "&&45ogscomments.linked"

Select «can be added up» contains "Y"
Field "parent_code"
GroupUp

Field TotalPoundsToSell
FormulaFill «unit conversion»*«45available»
Total

CollapseToLevel "1"
Field "TotalUnitsToSell"
FormulaFill TotalPoundsToSell

CollapseToLevel "Data"

Field "TotalPoundsToSell"
Fill ""
CollapseToLevel "1"

FormulaFill TotalUnitsToSell

CollapseToLevel "Data"
Field "TotalPoundsToSell"
PropagateUP

RemoveSummaries "7"

Field "TotalUnitsToSell"
FormulaFill TotalPoundsToSell/?(«unit conversion»>0,«unit conversion»,1)

Select «can be added up» contains "N"
Field "TotalUnitsToSell"
FormulaFill «45available»

EndNoShow

window "45ogscomments.linked"

Field "TotalUnitsToSell"
Select «» <> lookupselected("45ogscomments","IDNumber",IDNumber,"TotalUnitsToSell",0,0)
FormulaFill lookup("45ogscomments","IDNumber",IDNumber,"TotalUnitsToSell",0,0)

Field "TotalPoundsToSell"
Select «» <> lookupselected("45ogscomments","IDNumber",IDNumber,"TotalPoundsToSell",0,0)
FormulaFill lookup("45ogscomments","IDNumber",IDNumber,"TotalPoundsToSell",0,0)

SelectAll

Field "Item"
Sortup

message "Finished with Add-it-up!"
___ ENDPROCEDURE Add It Up _____________________________________________________

___ PROCEDURE Archive Item Search ______________________________________________
;Search for items that may be able to be moved to comments.archive

Field "Item"
Select «Listed» = |||No||| and «add/drop» = |||historical||| and «45available» <= 0 and «45inhouse» <= 0 

if info("empty")
    message "No items that need to be archived"
endif
Selectwithin «45soldtally» > «45tallyfilled»

if info("empty")
    message "No outstanding orders"
 endif
 
 message "Next Step- Run Archive Selected Items macro"







___ ENDPROCEDURE Archive Item Search ___________________________________________

___ PROCEDURE Archive Selected Items ___________________________________________
okcancel "Careful! Make sure you have only the items you want appended selected. Continue?"
    if clipboard() = "Cancel"
        stop
    endif
 openfile "Comments.archive"
 openfile "++45ogscomments.linked" 
 window "45ogscomments.linked"

;Delete appended records from Comments.Linked

noyes "Do you wish to delete appended items?"  

if clipboard() = "Yes"
   Loop
   Lastrecord
   Deleterecord
   until info("selected") = 1
endif






___ ENDPROCEDURE Archive Selected Items ________________________________________

___ PROCEDURE Create Cycle Count _______________________________________________
local vSection

vSection = ""

getscrap "Which section do you want to count? S### or F###"

select «Location1» contains clipboard()

field «Location1» sortup

vSection = clipboard()

openform "Cycle Count Form"

print dialog
___ ENDPROCEDURE Create Cycle Count ____________________________________________

___ PROCEDURE Export Inventory for Web _________________________________________
local waswindow, myfilename
local temp_array, sorted_array, lastorderno, theText
waswindow = info("windowname")

;; This macro produces a CSV of OGS's current inventory numbers that can be uploaded to the website.
;; It is up to OGS how often to run this. I think once a week would be sufficient, but it could even be run 
;; daily if you want!

;; Tracking Preference in the tally - exact wording matters. All sizes for a given item MUST have the same
;; tracking preference (or there may be unpredictable results). Recognized values are:

;; "Allow repack, no packup" (soil amendments, seed)
;; "Do not track inventory" (things we never run out of, like soil tests, mixing, and drop ship items)
;; "Track by item size" (tools, books - items that are not repackable)
;; "Track by item size, allowing repack" ("no holds barred repack" like "repacking" jiffy pots into cases and vice versa)

;; This macro depends on the "unit conversion" field being filled out and accurate for any items with a
;; "Track by item size, allowing repack" or "Allow repack, no packup" tracking preference.
;; Also, "Web Inventory Units" must be filled out for all items with one of the "repack" tracking preferences
;; and "Web Inventory Units" MUST be the same for all sizes of a given item.

;; This macro uses two unlinked files: "OGS web inventory" and "ogs_undownloaded."

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;; uses sparenumber3 to store a unit conversion on the bulk unlisted inventory

select Listed contains "bulk" and «Tracking Preference» contains "repack" and (sparenumber3 <> «45available»*«unit conversion»)
if info("empty")
else
    field sparenumber3
    formulafill «45available»*«unit conversion»
endif

;; excludes items with Listed="No" but includes items with Listed="bulk not listed"
select Listed <> "No"
field Item
Sortup
field parent_code
Groupup

;; this applies the last size code's tracking preference to the whole item, so if there are discrepancies, they get ignored
field «Tracking Preference»
Maximum

field «Web Inventory Units»
Maximum
field sparenumber3
Total
Outlinelevel 1

;; opens the unlinked file and replaces its contents with an export from ogscomments.linked
Openfile "OGS web inventory"
Openfile "&&45ogscomments.linked"
removedetail 0
removeallsummaries

field «inventory bulk»
formulafill lookup("45ogscomments.linked","parent_code",«parent_code»,"sparenumber3",zeroblank(0),1)

window waswindow
removeallsummaries

;; use formulafill to bring values from comments.linked into the unlinked file. 
;; Low stock threshold, out of stock threshold, 45available, and unit conversion are all brought over.

select Listed contains "web" and Item contains "A"
window "OGS web inventory"
field «low A»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos A»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory A»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «A unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "B"
window "OGS web inventory"
field «low B»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos B»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory B»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «B unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "C"
window "OGS web inventory"
field «low C»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos C»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory C»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «C unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "D"
window "OGS web inventory"
field «low D»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos D»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory D»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «D unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "E"
window "OGS web inventory"
field «low E»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos E»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory E»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «E unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "F"
window "OGS web inventory"
field «low F»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos F»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory F»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «F unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "G"
window "OGS web inventory"
field «low G»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos G»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory G»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «G unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "H"
window "OGS web inventory"
field «low H»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos H»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory H»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «H unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

window waswindow
select Listed contains "web" and Item contains "I"
window "OGS web inventory"
field «low I»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Low Stock Threshold",zeroblank(0),0)
field «oos I»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"Out of Stock Threshold",zeroblank(0),0)
field «inventory I»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"45available",zeroblank(0),0)
field «I unit conversion»
formulafill lookupselected("45ogscomments.linked","parent_code",«parent_code»,"unit conversion",zeroblank(0),0)

;; if there are any items with tracking preference 'N/A' assume they are not relevan to the web inventory upload
;; and remove them from the unlinked file
select «Tracking Preference» <> 'N/A'
removeunselected

;; replace tracking preference with values that the website will recognize
field «Tracking Preference»
formulafill replace(«Tracking Preference»,'Allow repack, no packup','repackNoPackup')
formulafill replace(«Tracking Preference»,'Do not track inventory','none')
formulafill replace(«Tracking Preference»,'Track by item size, allowing repack','repack')
formulafill replace(«Tracking Preference»,'Track by item size','size')

save

;;;;;;;;;;;;;;;;;;;;;;;;; get undownloaded orders from website ;;;;;;;;;;;;;;;;;;;;;;

;; This bit is important because it will account for any orders that have been placed online in the time since
;; the most recent orders in the tally were imported. Incorporating these undownloaded orders into the 
;; inventory calculation ensures that the inventory numbers we're uploading to the website are as accurate
;; as possible.

;; trying to open the tally without running the .Initialize macro
opensecret "45ogstally"
window "45ogstally:secret"

;;window "45ogstally"

;; The date restriction is just to return a smaller chunk of orders that need to be sorted, so that the macro runs more quickly.
;; Trying to find the most recent internet order that was placed. The group orders make it a little complicated, so that's
;; what this chunk of code is about.

select OrderNo > 320000 and OrderNo < 400000 and OrderPlaced >= (today()-7) and OrderNo = int(OrderNo)
if info("empty")
    ;; If there are no OGS internet orders in the tally, this will get all of the orders (will only be relevant 
    ;; for a short stretch of time after the fiscal year changeover but before the first batch of OGS orders is in the tally)
    LoadURL theText, "https://fedcoseeds.com/reports/ogs/collation?format=item"
else
    arrayselectedbuild temp_array, ¶, "",?(OriginalOrderNumber=0,str(OrderNo),str(OriginalOrderNumber))+¬+str(OrderNo)

    ;; sort by original order number
    sorted_array = arraymultisort(temp_array,¶,¬,"1n")

    ;; get OrderNo from last line in array
    lastorderno = arraylast(arraylast(sorted_array,¶),¬)
    
    ;;LoadURL theText, "https://landscape.fedcoseeds.com/api/v1/reports/ogs/collation?format=item&since-key=order&since-value=" + str(lastorderno)
    ;;LoadURL theText, "https://landscape.fedcoseeds.com/api/v1/reports/moose/collation"
    ;;LoadURL theText, "https://localhost/api/v1/reports/ogs/collation?since-key=order&format=item&since-value=" + "322478"
    Curl theText, "-k 'https://landscape.fedcoseeds.com/api/v1/reports/ogs/collation?since-key=order&format=item&since-value=" + str(lastorderno) + "'"

endif

debug

OpenFile "ogs_undownloaded"
OpenFile "&@theText"

window "OGS web inventory"

field undownloadedA
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyA",0,0)
field undownloadedB
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyB",0,0)
field undownloadedC
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyC",0,0)
field undownloadedD
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyD",0,0)
field undownloadedE
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyE",0,0)
field undownloadedF
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyF",0,0)
field undownloadedG
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyG",0,0)
field undownloadedH
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyH",0,0)
field undownloadedI
formulafill lookup("ogs_undownloaded","Item",«parent_code»,"qtyI",0,0)

field «inventory A»
formulafill zeroblank(«inventory A» - undownloadedA)
field «inventory B»
formulafill zeroblank(«inventory B» - undownloadedB)
field «inventory C»
formulafill zeroblank(«inventory C» - undownloadedC)
field «inventory D»
formulafill zeroblank(«inventory D» - undownloadedD)
field «inventory E»
formulafill zeroblank(«inventory E» - undownloadedE)
field «inventory F»
formulafill zeroblank(«inventory F» - undownloadedF)
field «inventory G»
formulafill zeroblank(«inventory G» - undownloadedG)
field «inventory H»
formulafill zeroblank(«inventory H» - undownloadedH)
field «inventory I»
formulafill zeroblank(«inventory I» - undownloadedI)

save

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

myfilename = "OGSInventory.csv"
;;export myfilename, exportline()+¶

export myfilename, str(parent_code)+","+«Tracking Preference»
+","+zbpattern(«A unit conversion»,"#.#####")+","+zbpattern(«B unit conversion»,"#.#####")+","+zbpattern(«C unit conversion»,"#.#")+","+zbpattern(«D unit conversion»,"#.#")+","+zbpattern(«E unit conversion»,"#.#")+","+zbpattern(«F unit conversion»,"#.#")+","+zbpattern(«G unit conversion»,"#.#")+","+zbpattern(«H unit conversion»,"#.#")+","+zbpattern(«I unit conversion»,"#.#")+","+«Web Inventory Units»
+","+zbpattern(«low A»,"#")+","+zbpattern(«low B»,"#")+","+zbpattern(«low C»,"#")+","+zbpattern(«low D»,"#")+","+zbpattern(«low E»,"#")+","+zbpattern(«low F»,"#")+","+zbpattern(«low G»,"#")+","+zbpattern(«low H»,"#")+","+zbpattern(«low I»,"#")+","+zbpattern(«low bulk»,"#")
+","+zbpattern(«oos A»,"#")+","+zbpattern(«oos B»,"#")+","+zbpattern(«oos C»,"#")+","+zbpattern(«oos D»,"#.#")+","+zbpattern(«oos E»,"#.#")+","+zbpattern(«oos F»,"#.#")+","+zbpattern(«oos G»,"#.#")+","+zbpattern(«oos H»,"#.#")+","+zbpattern(«oos I»,"#.#")+","+zbpattern(«oos bulk»,"#")
+","+zbpattern(«inventory A»,"#")+","+zbpattern(«inventory B»,"#")+","+zbpattern(«inventory C»,"#")+","+zbpattern(«inventory D»,"#")+","+zbpattern(«inventory E»,"#")+","+zbpattern(«inventory F»,"#")+","+zbpattern(«inventory G»,"#")+","+zbpattern(«inventory H»,"#")+","+zbpattern(«inventory I»,"#")+","+zbpattern(«inventory bulk»,"#")
+¶

;;; edit to blank out tracking preference after the first upload?

message "Done! Remember to upload 'OGSInventory.csv' to the website."

shellopendocument "https://fedcoseeds.com/manage_site/inventory/ogs"

//TEST SITE BELOW
//https://sarah.fedcoseeds.com/manage_site/inventory/ogs

___ ENDPROCEDURE Export Inventory for Web ______________________________________

___ PROCEDURE Hide Most Fields _________________________________________________
local fieldstoshow
fieldstoshow = "Listed"+¶+"IDNumber"+¶+"Item"+¶+"Description"+¶+"Sz."+¶+"ActWt"+¶+"Comments"+¶+"notes"+¶+"45available"+¶+"TotalPoundsToSell"+¶+"TotalUnitsToSell"+¶+"Initial"+¶+"Adjustments"+¶+"45sold"+¶+"45inhouse"+¶+"45tallyfilled"+¶+"45soldtally"+¶+"45soldwalkin"+¶+"45repack"+¶+"unit note"

showthesefields fieldstoshow
___ ENDPROCEDURE Hide Most Fields ______________________________________________

___ PROCEDURE Inquiry/I ________________________________________________________
field Item
sortup
openform "Inquiry"



___ ENDPROCEDURE Inquiry/I _____________________________________________________

___ PROCEDURE Import On Hold Items _____________________________________________
rememberwindow
openfile "45orderedogs"

originalwindow

field «OnHold»
select «» <> lookupselected("45orderedogs","IDNumber", IDNumber, "fill", 0,1)

Selectwithin IDNumber=lookupselected("45orderedogs","IDNumber",IDNumber,"IDNumber",0,1)

field «OnHold»
formulafill lookupselected("45orderedogs","IDNumber", IDNumber, "fill", 0,1)

window "45orderedogs"
closewindow

originalwindow

message "On Hold import complete"
___ ENDPROCEDURE Import On Hold Items __________________________________________

___ PROCEDURE Item History _____________________________________________________
Openform "Item History"
___ ENDPROCEDURE Item History __________________________________________________

___ PROCEDURE Load Repack Cost _________________________________________________
local vItemNum, vSupplyCost, wCL, wRepack

vItemNum = ""
vSupplyCost = ""
wCL = info("windowname")
wRepack = ""

openfile "45ogsrepack"
wRepack = info("windowname")

window wCL

selectall

field «Item»
sortup

Select sizeof("Type of Repack") > 0

firstrecord

noshow

loop

vItemNum = «Item»

window wRepack
selectall
select ItemNumMade1 = vItemNum

case
    info("selected") < info("records")
    field «Date Finished»
    sortup
    field «RepackCostPerUnit»
    lastrecord
    vSupplyCost = «RepackCostPerUnit»
    window wCL
    bag = vSupplyCost 
case
    info("selected") = info("records")
    window wCL
endcase

downrecord

until info("stopped")

endnoshow

message "Done!"


    

___ ENDPROCEDURE Load Repack Cost ______________________________________________

___ PROCEDURE Negative Report __________________________________________________
;; searches for items that are less than 0 for Stasha to go and investigate.

global vTitle, vWindowlist

vTitle = ""
vWindowlist = ""

gosheet

vWindowlist = info("windows")

/* This code below determines if the Inventory Form is already open and closes
that form if it is. This helps to refresh the title in the form*/

if
    vWindowlist contains "Inventory Form"
        window "45ogscomments.linked:Inventory Form"
        closewindow
endif

Field Item
    sortup
    
Field "Item" Select «45available» < 0 or «45inhouse» < 0

vTitle = "Negative Report" + ¬ + datepattern(today(), "Month ddnth, yyyy")

openform "Inventory Form"

Print Dialog

___ ENDPROCEDURE Negative Report _______________________________________________

___ PROCEDURE Open Inventory Summary ___________________________________________
openform "Sara Form"
___ ENDPROCEDURE Open Inventory Summary ________________________________________

___ PROCEDURE Open Repack ______________________________________________________
Openfile "45ogsrepack"
gosheet
___ ENDPROCEDURE Open Repack ___________________________________________________

___ PROCEDURE Priceline ________________________________________________________
Field "add/drop"
Select «priceline» notcontains |||21||| or sizeof("priceline") = 0 and («add/drop» notcontains |||drop|||) 
___ ENDPROCEDURE Priceline _____________________________________________________

___ PROCEDURE Repack Pending Report ____________________________________________
;;Searches for items that have have "Repack Pending" in the notes.

Field "Item"
sortup

Local fieldstoshow

fieldstoshow = "Item"+¶+"Description"+¶+"Sz."+¶+"Comments"+¶+"Notes"+¶+"45available"

Field "Item"

Select notes contains "Repack"

openform "Inventory Report"

___ ENDPROCEDURE Repack Pending Report _________________________________________

___ PROCEDURE Sale History _____________________________________________________
;;Searches sale history

Local fieldstoshow
fieldstoshow = "Item"+¶+"Description"+¶+"45available"+¶+"45sold"+¶+"44sold"+¶+"43sold"+¶+"42sold"+¶+"41sold"

showthesefields fieldstoshow

___ ENDPROCEDURE Sale History __________________________________________________

___ PROCEDURE Update Repack Numbers ____________________________________________
rememberwindow
openfile "repack_vertical_file"

originalwindow

Field ("45repack")
Select «» <> lookupselected("repack_vertical_file","IDNumber",IDNumber,"Net",0,0)

selectwithin IDNumber=lookupselected("repack_vertical_file","IDNumber",IDNumber,"IDNumber",0,0)

Field ("45repack")
formulafill lookupselected("repack_vertical_file","IDNumber",IDNumber,"Net",0,0)

call "tabdown/2"

;SelectAll
window "repack_vertical_file"
closewindow

originalwindow

message "Finished Repack Update"
___ ENDPROCEDURE Update Repack Numbers _________________________________________

___ PROCEDURE Update Tally Numbers _____________________________________________
rememberwindow
openfile "45orderedogs"

originalwindow

field ("45soldtally")
select «» <> lookupselected("45orderedogs","IDNumber", IDNumber, "qty", 0,1)

Selectwithin IDNumber=lookupselected("45orderedogs","IDNumber",IDNumber,"IDNumber",0,1)

field ("45soldtally")
formulafill lookupselected("45orderedogs","IDNumber", IDNumber, "qty", 0,1)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Field ("45tallyfilled")
Select «» <> lookupselected("45orderedogs","IDNumber",IDNumber,"fill",0,1)

selectwithin IDNumber=lookupselected("45orderedogs","IDNumber",IDNumber,"IDNumber",0,1)

Field ("45tallyfilled")
formulafill lookupselected("45orderedogs","IDNumber",IDNumber,"fill",0,1)

;;SelectAll

window "45orderedogs"
closewindow

originalwindow

message "Finished Tally Sold/Filled Update"
___ ENDPROCEDURE Update Tally Numbers __________________________________________

___ PROCEDURE Update Walk-In Numbers ___________________________________________
rememberwindow

;;--------------------------- transfers -----------------------;;
yesno "Any Transfers?"
if clipboard() contains "Yes"

openfile "45transfers_vertical_ogs"

originalwindow
Field ("45transfers")
Select «» <> lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

selectwithin IDNumber=lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"IDNumber",0,1)

Field ("45transfers")
formulafill lookupselected("45transfers_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

window "45transfers_vertical_ogs"
closewindow

originalwindow

field «45transfers»
call "tabdown/2"

else
endif
SelectAll

openfile "45walkin_vertical_ogs"

;;--------------------------- normal sales -----------------------;;

originalwindow
Field ("45soldwalkin")
Select «» <> lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

selectwithin IDNumber=lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"IDNumber",0,1)

Field ("45soldwalkin")
formulafill lookupselected("45walkin_vertical_ogs","IDNumber",IDNumber,"qty",0,1)

;;SelectAll
window "45walkin_vertical_ogs"
closewindow

originalwindow

field «45soldwalkin»
call "tabdown/2"

message "Finished Walk-in Sales Update"



___ ENDPROCEDURE Update Walk-In Numbers ________________________________________

___ PROCEDURE Zero Audit Form __________________________________________________
Openform "Zero Audit Form"
Message "When printing make sure to select 1 from 1 for print options, NOT all"
Print dialog

___ ENDPROCEDURE Zero Audit Form _______________________________________________

___ PROCEDURE Zero Report ______________________________________________________
;;Searches for items that have a 0 "45Available" to make sure that comments are filled in and counts are accurate.

global vTitle, vWindowlist

vTitle = ""
vWindowlist = ""

gosheet

vWindowlist = info("windows")

/* This code below determines if the Inventory Form is already open and closes
that form if it is. This helps to refresh the title in the form*/

if
    vWindowlist contains "Inventory Form"
        window "45ogscomments.linked:Inventory Form"
        closewindow
endif

Field "Item"
sortup

Select («45available» = 0 or «45inhouse» = 0) and «Comments» notcontains "clearance" and «notes» notcontains "clearance" and «Comments» notcontains "service"

yesno "Only show items with no inventory and no comments?"
    if 
        clipboard() contains "yes"
            selectwithin «Comments» = ""
    endif
    
vTitle = "Zero Report" + ¬ + datepattern(today(), "Month ddnth, yyyy")

openform "Inventory Form"

print dialog

___ ENDPROCEDURE Zero Report ___________________________________________________

___ PROCEDURE (Labels) _________________________________________________________

___ ENDPROCEDURE (Labels) ______________________________________________________

___ PROCEDURE OGS Product Label ________________________________________________
Openform "OGS Product Label"
___ ENDPROCEDURE OGS Product Label _____________________________________________

___ PROCEDURE OGS Line Labels Big Bags _________________________________________
Openform "OGS Line Labels Big Bags"
___ ENDPROCEDURE OGS Line Labels Big Bags ______________________________________

___ PROCEDURE OGS Line Label ___________________________________________________
Openform "OGS Line Label"
___ ENDPROCEDURE OGS Line Label ________________________________________________

___ PROCEDURE cost/5 ___________________________________________________________
field «45cost»
___ ENDPROCEDURE cost/5 ________________________________________________________

___ PROCEDURE Avail/8 __________________________________________________________
field «45available»
___ ENDPROCEDURE Avail/8 _______________________________________________________

___ PROCEDURE cost/6 ___________________________________________________________
field «45cost»
___ ENDPROCEDURE cost/6 ________________________________________________________

___ PROCEDURE avail/8 __________________________________________________________
field «45available»
___ ENDPROCEDURE avail/8 _______________________________________________________

___ PROCEDURE (CommonFunctions) ________________________________________________

___ ENDPROCEDURE (CommonFunctions) _____________________________________________

___ PROCEDURE ExportMacros _____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros are saved to your clipboard!"
___ ENDPROCEDURE ExportMacros __________________________________________________

___ PROCEDURE ImportMacros _____________________________________________________
local Dictionary1,Dictionary2, ProcedureList
Dictionary1=""
Dictionary1=clipboard()
yesno "Press yes to import all macros from clipboard"
if clipboard()="No"
stop
endif
//step one
importdictprocedures Dictionary1, Dictionary2
//changes the easy to read macros into a panorama readable file

 
//step 2
//this lets you load your changes back in from an editor and put them in
//copy your changed full procedure list back to your clipboard
//now comment out from step one to step 2
//run the procedure one step at a time to load the new list on your clipboard back in
//Dictionary2=clipboard()
loadallprocedures Dictionary2,ProcedureList
message ProcedureList //messages which procedures got changed

___ ENDPROCEDURE ImportMacros __________________________________________________

___ PROCEDURE Symbol Reference _________________________________________________
bigmessage "Option+7= ¶  [in some functions use chr(13)
Option+= ≠ [not equal to]
Option+\= « || Option+Shift+\= » [chevron]
Option+L= ¬ [tab]
Option+Z= Ω [lineitem or Omega]
Option+V= √ [checkmark]
Option+M= µ [nano]
Option+<or>= ≤or≥ [than or equal to]"


___ ENDPROCEDURE Symbol Reference ______________________________________________

___ PROCEDURE GetDBInfo ________________________________________________________
local DBChoice, vAnswer1, vClipHold

Message "This Procedure will give you the names of Fields, procedures, etc in the Database"
//The spaces are to make it look nicer on the text box
DBChoice="fields
forms
procedures
permanent
folder
level
autosave
fileglobals
filevariables
fieldtypes
records
selected
changes"
superchoicedialog DBChoice,vAnswer1,“caption="What Info Would You Like?"
captionheight=1”


vClipHold=dbinfo(vAnswer1,"")
bigmessage "Your clipboard now has the name(s) of "+str(vAnswer1)+"(s)"+¶+
"Preview: "+¶+str(vClipHold)
Clipboard()=vClipHold

___ ENDPROCEDURE GetDBInfo _____________________________________________________

___ PROCEDURE .AutomaticFY _____________________________________________________

    
global dateHold, dateMath, intYear, 
thisFYear,lastFYear,nextFYear,intMonth,fileDate

fileDate=val(striptonum(info("databasename")))
nextFYear=""
thisFYear=""
lastFYear=""

//get the date
dateHold = datepattern(today(),"mm/yyyy")

//gets the current month and year
intMonth = val(dateHold[1,"/"][1,-2])
intYear = val(dateHold["/",-1][2,-1])

//assigns FY numbers for years

case val(intMonth)>6
    nextFYear=str(intYear-1976)
    thisFYear=str(intYear-1977)
    lastFYear=str(intYear-1978)

case val(intMonth)<7
    nextFYear=str(intYear-1977)
    thisFYear=str(intYear-1978)
    lastFYear=str(intYear-1979)

endcase

//checks if this is an older file and needs older FYs
if fileDate ≤ val(lastFYear) and fileDate > 0
    nextFYear=str(fileDate+1)
    thisFYear=str(fileDate)
    lastFYear=str(fileDate-1)
endif

//tallmessage str(nextFYear)+¬+str(thisFYear)+¬+str(lastFYear)


/*

///////~~~~~~~
Programmer Notes
~~~~~~~~~//////////
The danger of this procedure is that come July 1st of the year, it will automatically set
to open the newest files of a non-numbered Panorama file. And if those don't exist, you're 
gonna see errors. Also, a non numbered Panorama file that needs to call older files shouldn't
use this macro



To use these variables please note the following Panorama syntax rules:


filenames using variables:
    can just concatenate as a string
    
    ex:
        
    openfile str(variable)+"filename" 


field calls using variables:
    best to be only one variable and nothing else
    must be surrounded by ( )
    
    ex:
    
    field (VariableFieldName)
    
do your math and/or concatenation into the variable before calling it
    VariableFieldName=str(variable)+"fieldname"
 
field (str(variable)+"fieldname") will work but can cause errors
    
for assignments to that variable'd field 
    use «» for "current field/current cell" 
    
    ex: 
   
    «» = "10"
  
    
*/

___ ENDPROCEDURE .AutomaticFY __________________________________________________

___ PROCEDURE Folders&FilesMacros ______________________________________________

//message "This Function is meant to get you information about the folders and path your files are in for Panorama"


global commList, commWanted, clipHoldComm, buttonChoice, numChoice

commList=""
commWanted=""
clipHoldComm=""
buttonChoice=""
numChoice=0

commList=¶+
    "1 - Copy Text of folderpath"
    +¶+¬+¬+¬+¬+¬+¬+
    "1 code -- folderpath(folder(""))"
    +¶+" "+¶+
    "2 - Copy list of All Files and Folders in this folder" 
    +¶+¬+¬+¬+¬+¬+¬+
    "2 code -- listfiles(folder(""),"")"
    +¶+" "+¶+
    "3 - Copy list of All Panorama files in this folder" 
    +¶+¬+¬+¬+¬+¬+¬+
    '3 code -- listfiles(folder(""),"????KASX")'
    +¶+" "+¶+
    "4 - Copy list of All Text files in this folder" 
    +¶+¬+¬+¬+¬+¬+¬+
    '4 code -- listfiles(folder(""),"TEXT????")'

/*

//NOTE: these quotation marks “” vs "" are called smart quotes
//you get them with opt+[ and opt+shift+[
//normally for superchoicedialogs, i would use curly brackets around title or caption
//but to have this be able to be written into new files from another macro, I had
//to use smart quotes

*/
superchoicedialog commList, commWanted, 
“Title="Get File/Folder/Path"
    Caption="1 - Copy ~~~~~~ gets you the data
        1 - Code ~~~~~~ gets you the formula"
    captionheight="2"
    buttons="Ok;Cancel"
    width="800"
    height="800"”
    

        clipHoldComm=commWanted
        numChoice=striptonum(clipHoldComm)[1,3]


if commWanted[1,12] notcontains "code"

    case numChoice="1"
        tallmessage "clipboard now has: "+¶+folderpath(folder(""))
        clipboard()=folderpath(folder(""))

    case numChoice="2"
        tallmessage "clipboard now has: "+¶+listfiles(folder(""),"")
        clipboard()=listfiles(folder(""),"")
    
    case numChoice="3"
        tallmessage "clipboard now has: "+¶+listfiles(folder(""),"????KASX")
        clipboard()=listfiles(folder(""),"????KASX")

    case numChoice="4"
        tallmessage "clipboard now has: "+¶+listfiles(folder(""),"TEXT????")
        clipboard()=listfiles(folder(""),"TEXT????")

    endcase
endif

if commWanted[1,12] contains "code"
    case numChoice="1"
    clipboard()='folderpath(folder(""))'
    tallmessage "clipboard now has: "+¶+'folderpath(folder(""))'

    case numChoice="2"
    clipboard()='listfiles(folder(""),"")'
    tallmessage "clipboard now has: "+¶+'listfiles(folder(""),"")'
    
    case numChoice="3"
        tallmessage "clipboard now has: "+¶+'listfiles(folder(""),"????KASX")'
        clipboard()='listfiles(folder(""),"????KASX")'

    case numChoice="4"
        tallmessage "clipboard now has: "+¶+'listfiles(folder(""),"TEXT????")'
        clipboard()='listfiles(folder(""),"TEXT????")'

    endcase
endif
    


___ ENDPROCEDURE Folders&FilesMacros ___________________________________________

___ PROCEDURE DesignSheetExportImport __________________________________________

global vdictionary, 
name, value, ImportExportChoicelist,
fileList,choiceMade,winChoice1,winChoice2,vOptions

/*
programmer's notes

i was testing using a variable for options as stated in the reference. it seems to work with vOptions below

also tested using a call to listfiles vs putting listfiles in a variable. both seem to work

as seen in other procedures in this file instead of using curly braces we are using smartquotes because "setproceduretext" for the AddMacros fuction wont work otherwise

also, options for superchoices and other customizable dialogs are very particular in 
their syntax

caption = "dafsdf" will allow the code to run, but will not show a caption
caption="dafsdf" will actually show the caption
*/


vOptions=“caption="Choose file to export Design Sheet from"”
choiceMade=""
fileList=listfiles(folder(""),"????KASX")


superchoicedialog fileList, choiceMade, vOptions

winChoice1=choiceMade

superchoicedialog fileList, choiceMade,
“caption="Choose file to export Design Sheet to"”

winChoice2=choiceMade

window (winChoice1)
    opendesignsheet
    vdictionary=""
    firstrecord

        loop
            setdictionaryvalue vdictionary, «Field Name», «Equation»
            downrecord
        until info("stopped")

window (winChoice2)
    opendesignsheet
    firstrecord

        loop
            field «Equation»
            «» = getdictionaryvalue(vdictionary, «Field Name»)
            downrecord
        until info("stopped")


___ ENDPROCEDURE DesignSheetExportImport _______________________________________

___ PROCEDURE .FileChecker _____________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________
///________________________________This is the .FileChecker macro in GetMacros_________________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________


local fileNeeded,folderArray,smallFolderArray,sizeCheck,
procList,sizeCheck,procNames,procDBs,mostRecentProc

///________________________EDITME_____________
//replace this with whatever file you're error checking
//----------------------//
fileNeeded="members"    //
//----------------------//


////_____Got the file, but it's not open?_______________
case info("files") notcontains fileNeeded and listfiles(folder(""),"????KASX") contains fileNeeded
openfile fileNeeded

///________Don't got the file?__________________
case listfiles(folder(""),"????KASX") notcontains fileNeeded


    procList=arraystrip(info("procedurestack"),¶)
    sizeCheck=arraysize(procList,¶)
        if sizeCheck>1
            procList=arrayrange(procList,2,sizeCheck,¶) //this is to exclude getting recursive info about this macro, especially while testing
        else
            procList=arraystrip(info("procedurestack"),¶)
        endif

    procNames=arraycolumn(procList,1,¶,¬)
    procDBs=arraycolumn(procList,2,¶,¬)
    mostRecentProc=array(procNames,1,¶) 
    folderArray=folderpath(folder(""))
    sizeCheck=arraysize(folderArray,":")
    smallFolderArray=arrayrange(folderArray,4,sizeCheck,":")

displaydata "Error:"
+¶+
"You are missing the '"+fileNeeded+
"' Panorama file in this folder 
and can't continue the '"+mostRecentProc+"' procedure without it. 
Please move a copy of '"+fileNeeded+
"' to the appropriate folder and try the procedure again"
+¶+¶+¶+
"folder you're currently running from is: "
+¶+
smallFolderArray
+¶+¶+¶+
"current Pan files in that folder are: "
+¶+
listfiles(folder(""),"????KASX")
+¶+¶+¶+
"Pressing 'Ok' will open the Finder to your current folder"
+¶+¶+
"Press 'Stop' will stop this procedure", “title="Missing File!!!!" captionwidth=900 size=17 height=500 width=800”
//___________________________
//note, the above are "smart quotes" option+[ and option+shift+[ 
//you can also use curley braces, but another program I run will break
//if this file has thos
//___________________________

revealinfinder folder(""),""
stop

///_______File is open, but not active?______
defaultcase
window fileNeeded

endcase

/*
Example:

You are missing the 'members' Panorama file in this folder 
and can't continue this procedure without it. Please move a copy of
'members' to the appropriate folder and try the procedure again


folder you're currently running from is: 
Desktop:Panorama:FY45 Panorama Projects:GetMacros:


current Pan files in that folder are: 
GetMacros
GetMacrosDL
GetMacros44


Pressing 'Ok' will open the Finder to your current folder

Press 'Stop' will stop this procedure
*/
___ ENDPROCEDURE .FileChecker __________________________________________________

___ PROCEDURE .GetErrorLog _____________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________
///________________________________This is the .GetErrorLog macro in GetMacros_________________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________
/*

This can be called to with a parameter of the 
info("error") statement to display the error, give
the user the opportunity to try again or continue 
despite the error.

Either way, it makes a log of the error and what procedures,
windows, files, and variables were in use. 

-Lunar 8-22

Syntax to call:

        if error
            call .GetErrorLog,info("error")
        endif

*/
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________

fileglobal fileNeeded,folderArray,smallFolderArray,sizeCheck, procList, mostRecentProc, 
panFilesList,activeFiles,allvariables,procNames,procDBs,errorList, procText, procTextArray,
lineNum, procCount, usedvariables,printVariables,strippedText,getError,errorMsg,vDb,vProc,
activeWindows,DictNameToday

//this is to keep a log of the errors
permanent errorDictionary
    errorDictionary=errorDictionary
    if error
    errorDictionary=""
    endif

errorMsg=""

getError=str(parameter(1))
    if error //if there's no parameter given, or if info("error") is blank, then say "Unknown"
    getError="Unknown"
    endif

procList=arraystrip(info("procedurestack"),¶)
    if procList="" //sometimes, there's no info in the procedure stack, and this macro shoudl stop at this point
    message "Procedure Stack is Empty -L"
    stop
    endif
sizeCheck=arraysize(procList,¶)
    if sizeCheck>1
    procList=arrayrange(procList,2,sizeCheck,¶) //this is to exclude getting recursive info about this macro, especially while testing
    else
    procList=arraystrip(info("procedurestack"),¶)
    endif

procNames=arraycolumn(procList,1,¶,¬)
procDBs=arraycolumn(procList,2,¶,¬)
mostRecentProc=array(procNames,1,¶) 
folderArray=folderpath(folder(""))

///____________more readable filepath________________
;sizeCheck=arraysize(folderArray,":")
;smallFolderArray=arrayrange(folderArray,4,sizeCheck,":")
///__________________________________________________

panFilesList=listfiles(folder(""),"????KASX")
activeFiles=info("files")
activeWindows=info("windows")
allvariables="Global variables"+¶+¶+info("globalvariables")+¶+¶+"local variables"+¶+¶+info("localvariables")+¶+¶+"fileglobal variables"+¶+¶+info("filevariables")+¶+¶+"window variables"+¶+¶+info("windowvariables")

//____bugcheck_______
;displaydata procNames
;displaydata procDBs
//___________________

lineNum=1
procCount=arraysize(procNames,¶)
procTextArray=""

//_______build an array of the procedure text of all the last used procedures_____
/*
Notes: this kept breaking when I tried to use arrayfilters or arraybuilds, and apparently there's 
known issues with using local variables that throws an exception about the call() procedure
because its using a subroutine using EXECTUTE to do arrayfilters 

That being said, it kept breaking even after turning the variables global, so now it's a loop
*/
//________________________________________________________________________________

loop
    vDb=array(procDBs,lineNum,¶)
    vProc=array(procNames,lineNum,¶)
        getproceduretext vDb,vProc,procText
        procTextArray=vProc+¶+¶+procText+¶+procTextArray  //format: Name of Procedure, two returns, text from the proc, then the last thing added put on the end
    lineNum=lineNum+1
while lineNum<procCount


//_________________Make code into word array______________________//
/*
this function makes two arrays similar enough to compare to find out
which of the active variables was in the procedures that were recently called
gets rid of the most common characters in the text and replaces them with ; to give the other functions
a separator to work with

This was done because there's like 30 variables that are only Panorama's that also gets included in the INFO("xVARIABLE")
calls, and those aren't really useful for bugfixing procedures. 

//______________get out extra characters, but retrain spaces between words using a semicolon__________
note: there were also pipes '||' with curley brackets between them, but those break the SETPROCEDURETEXT statement, so I had to take them out of the "GetMacros" version
*/
strippedText=replacemultiple(procTextArray,
“.||?||!||,||;||:||-||_||(||)||[||]||"||'||+||¶||¬||/||=||*||" "|| ||”,
“;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||”,
"||")

strippedText=stripchar(strippedText,"AZaz09;")
arraystrip strippedText,";"

//_____Change the format of the array into a ¶ one_______
strippedText=replace(strippedText,";",¶)
arraydeduplicate strippedText,strippedText,¶

//________get variablelist into a cleaner version_____
usedvariables=arraystrip(allvariables,¶)

//__________do a comparison for whats in both of them and put that in printVariables
arrayboth strippedText, usedvariables, ¶, printVariables

//_______Print Check_____
;displaydata printVariables


////____________Error Log____________________
/*
The short form of what gets displayed to the user specifically to be added to the 
errorDictionary. You can get this full log by calling 
DISPLAYDATA errorDictionary
*/
DictNameToday=superdatepattern(supernow(),"mm/dd/yy@", "hh:mm" )
setdictionaryvalue errorDictionary,DictNameToday, 
"Error: '"+mostRecentProc+"' created an error."
+¶+¶+
"ErrorCode: "+getError
+¶+¶+¶+
"folder in use: "
+¶+
folderArray
+¶+¶+¶+
"current Pan files in that folder are: "
+¶+
panFilesList
+¶+¶+¶+
"currently open files are: "
+¶+
activeFiles
+¶+¶+¶+
"currently open windows are: "
+¶+
activeWindows
+¶+¶+¶+
"last procedures run were"
+¶+
procList
+¶+¶+¶+
"text of non-design/form procedures:"
+¶+
procTextArray
+¶+¶+¶+
"variables used in last macros:"
+¶+
printVariables

///__________Future feature_____________
/*
Give the user instructions on what to do based on the error
*/
errorList="array of errors to give advice about"
//_______________________________________


////_____________ErrorDisplay for user________________________________________

displaydata "Error: '"+mostRecentProc+"' procedure/macro created an error."
+¶+¶+
"ErrorCode: "+getError
+¶+¶+
"Warning! If you click OK the macro will continue without fixing
the error. Proceed with caution, or click Stop instead."
+¶+¶+
"Click 'stop' to end the macro here and try what you were doing again"
+¶+¶+
"If the problem persists, use the 'COPY' button, paste this error in an e-mail 
and send it to: tech-support@fedcoseeds.com with a description of what happened



_______________________________________________________________________________"
+¶+¶+¶+
"---------------------------------------------------
THE FOLLOWING LINES ARE TO HELP WITH ERROR CHECKING
---------------------------------------------------"
+¶+¶+¶+
"folder in use: "
+¶+
folderArray
+¶+¶+¶+
"current Pan files in that folder are: "
+¶+
panFilesList
+¶+¶+¶+
"currently open files are: "
+¶+
activeFiles
+¶+¶+¶+
"currently open windows are: "
+¶+
activeWindows
+¶+¶+¶+
"last procedures run were"
+¶+
procList
+¶+¶+¶+
"text of non-design/form procedures:"
+¶+
procTextArray
+¶+¶+¶+
"variables used in last macros:"
+¶+
printVariables, 
“title="Error Capture Bot 3.0" 
captionwidth=900 
size=17 
height=500 
width=1000”



/*
//_________What this error looks like___________

Error: '.GetErrorLog' procedure/macro created an error.

ErrorCode: Unknown

Warning! If you click OK the macro will continue without fixing
the error. Proceed with caution, or click Stop instead.

Click 'stop' to end the macro here and try what you were doing again

If the problem persists, use the 'COPY' button, paste this error in an e-mail 
and send it to: tech-support@fedcoseeds.com with a description of what happened



_______________________________________________________________________________


---------------------------------------------------
THE FOLLOWING LINES ARE TO HELP WITH ERROR CHECKING
---------------------------------------------------


folder in use: 
LunarWindflower:Applications:Panorama:Panorama.app:Contents:MacOS:


current Pan files in that folder are: 
ProVUE Registration.pan


currently open files are: 
Untitled


currently open windows are: 
Untitled
Untitled:.GetErrorLog


last procedures run were
.GetErrorLog	Untitled		0


text of non-design/form procedures:
.GetErrorLog

The Text of the .GetErrorLog macro would be here, but I don't wanna double this file's
Length


variables used in last macros:
activeFiles
activeWindows
allvariables
DictNameToday
errorDictionary
errorList
errorMsg
fileNeeded
folderArray
getError
lineNum
mostRecentProc
panFilesList
printVariables
procCount
procDBs
procList
procNames
procText
procTextArray
sizeCheck
smallFolderArray
strippedText
usedvariables
vDb
vProc

*/

___ ENDPROCEDURE .GetErrorLog __________________________________________________

___ PROCEDURE SeeErrorLog ______________________________________________________

    displaydata errorDictionary

___ ENDPROCEDURE SeeErrorLog ___________________________________________________

___ PROCEDURE .WaitXSeconds ____________________________________________________
local start, end,secondsToWait

secondsToWait=5
start=now()
end=start+secondsToWait
loop
    nop
while now()≤end

//_____test timer____
;message end - start

___ ENDPROCEDURE .WaitXSeconds _________________________________________________

___ PROCEDURE GetWindowSize ____________________________________________________
global newrec, rectangle1,RecTop,RecLeft,RecHeight,RecWidth,whichWin,winList2

winList2=info("windows")
superchoicedialog winList2,whichWin,“caption="Which Window do you want the size of?"”
window (whichWin)
rectangle1=info("windowrectangle")
RecTop=rtop(rectangle1)
RecLeft=rleft(rectangle1)
RecHeight=rheight(rectangle1)
RecWidth=rwidth(rectangle1)

newrec=str(RecTop)+","+str(RecLeft)+","+str(RecHeight)+","+str(RecWidth)
message "You now have the Top, Left, Height, and Width of the window. You can use the setwindow command with these numbers"
clipboard()=newrec
//top,left,height,width


___ ENDPROCEDURE GetWindowSize _________________________________________________
