usingSystem;
usingSystem.Collections.Generic;
usingSystem.Linq;
usingSystem.Text;
usingSystem.Threading.Tasks;

namespaceGenrate_Excel
{















Reverse Charge Transaction(Y or N)
Posting charges for
Enc.ID (not for Treat.Series)
ECD ID(only for Treat.Series)
Service Provider Service ID
 Service Date
Service Time
 Service Time DST/ST
Stop Date
 Stop Time
Stop Time DST/ST
Quantity
Extended Price
 Duration (in minutes)
Performing HP ID
Performing HP ID Issuer
 Ordering HP ID
 Ordering HP ID Issuer
Referring HP ID
Referring HP ID Issuer
 Supervising HP ID
 Supervising HP ID Issuer
Override Proc.Code
 Proc.Mod. (1)
Proc.Mod. (2)
Proc.Mod. (3)
Proc.Mod. (4)
Override Rev.Code
 Override Charge Amt.
 Override Service Name
 Diag.Code (1)
Diag.Code (2)
Diag.Code (3)
Diag.Code (4)
Diag.Code (5)
Diag.Code (6)
Diag.Code (7)
Diag.Code (8)
Dose Quantity
 Referral Number
Authorization Number
 Cost
 ABN Status Date
 ABN Status Code
 National Drug Code
 Item Number
Model Number
 Taxonomy Code
Clinic Code
 Tooth Designation
Dent.Srfc.Code1
Dent.Srfc.Code2
Dent.Srfc.Code3
Dent.Srfc.Code4
Dent.Srfc.Code5
Tooth Status Code
Oral Cavity Code1
Oral Cavity Code2
Oral Cavity Code3
Oral Cavity Code4
Oral Cavity Code5
Placement Status Code
Prior Placement Date
Podiatry Last PCP Visit Date
Hearing And Vision Prescription Date
Vision Category Code
Vision Certification Condition Indicator
 Vision Condition Indicator Code1
Vision Condition Indicator Code2
 Vision Condition Indicator Code3
Vision Condition Indicator Code4
 Vision Condition Indicator Code5
Dme Certificate of Medical Necessity Transmission Code
Dme Certification Type Code
 Dme Duration(Months)
Dme Certification Revision Date
 Dme initial Certification Date
Dme Last Certification Date
 Dme Length of Medical Necessity Days
Dme Rental Price
Dme Purchase Price
Dme Frequency Code
Dme Certification Condition Indicator
 Dme Condition Indicator Code1
Dme Condition Indicator Code2
 Special Processing Code1
 Special Processing Code2
 Special Processing Code3
 Special Processing Code4
 Special Processing Code5
 Special Processing Code6
 Investigational Device Exempt No
Place Of Service Override
 Procedure Code Description Override
Enc.Provider
Enc.Location
Enc.Matching Strategy Code
National Drug Code Quantity
 RX Number
Unit Of Measure Code
 Auto Charge Rule Text
Auto Charge Rule Type Code
Service Price Rule Text
 Error Message

     "Reverse Charge Transaction (Y or N)","Posting charges for","Enc. ID (not for Treat. Series)","ECD ID (only for Treat. Series)","Service Provider Service ID","Service Date","Service Time","Service Time DST/ST","Stop Date","Stop Time","Stop Time DST/ST","Quantity","Extended Price","Duration (in minutes)","Performing HP ID","Performing HP ID Issuer","Ordering HP ID","Ordering HP ID Issuer","Referring HP ID","Referring HP ID Issuer","Supervising HP ID","Supervising HP ID Issuer","Override Proc. Code","Proc. Mod. (1)","Proc. Mod. (2)","Proc. Mod. (3)","Proc. Mod. (4)","Override Rev. Code","Override Charge Amt.","Override Service Name","Diag. Code (1)","Diag. Code (2)","Diag. Code (3)","Diag. Code (4)","Diag. Code (5)","Diag. Code (6)","Diag. Code (7)","Diag. Code (8)","Dose Quantity","Referral Number","Authorization Number","Cost","ABN Status Date","ABN Status Code","National Drug Code","Item Number","Model Number","Taxonomy Code","Clinic Code","Tooth Designation","Dent. Srfc. Code1","Dent. Srfc. Code2","Dent. Srfc. Code3","Dent. Srfc. Code4","Dent. Srfc. Code5","Tooth Status Code","Oral Cavity Code1","Oral Cavity Code2","Oral Cavity Code3","Oral Cavity Code4","Oral Cavity Code5","Placement Status Code","Prior Placement Date","Podiatry Last PCP Visit Date","Hearing And Vision Prescription Date","Vision Category Code","Vision Certification Condition Indicator","Vision Condition Indicator Code1","Vision Condition Indicator Code2","Vision Condition Indicator Code3","Vision Condition Indicator Code4","Vision Condition Indicator Code5","Dme Certificate of Medical Necessity Transmission Code","Dme Certification Type Code","Dme Duration(Months)","Dme Certification Revision Date","Dme initial Certification Date","Dme Last Certification Date","Dme Length of Medical Necessity Days","Dme Rental Price","Dme Purchase Price","Dme Frequency Code","Dme Certification Condition Indicator","Dme Condition Indicator Code1","Dme Condition Indicator Code2","Special Processing Code1","Special Processing Code2","Special Processing Code3","Special Processing Code4","Special Processing Code5","Special Processing Code6","Investigational Device Exempt No","Place Of Service Override","Procedure Code Description Override","Enc. Provider","Enc. Location","Enc. Matching Strategy Code","National Drug Code Quantity","RX Number","Unit Of Measure Code","Auto Charge Rule Text","Auto Charge Rule Type Code","Service Price Rule Text","Error Message"


      classClass1
{
table.Columns.Add("Reverse Charge Transaction (Y or N)"
typeof(string));
table.Columns.Add("Posting charges for"
typeof(string));
table.Columns.Add("Enc. ID (not for Treat. Series)"
typeof(string));
table.Columns.Add("ECD ID (only for Treat. Series)"
typeof(string));
table.Columns.Add("Service Provider Service ID"
typeof(string));
table.Columns.Add("Service Date"
typeof(string));
table.Columns.Add("Service Time"
typeof(string));
table.Columns.Add("Service Time DST/ST"
typeof(string));
table.Columns.Add("Stop Date"
typeof(string));
table.Columns.Add("Stop Time"
typeof(string));
table.Columns.Add("Stop Time DST/ST"
typeof(string));
table.Columns.Add("Quantity"
typeof(string));
table.Columns.Add("Extended Price"
typeof(string));
table.Columns.Add("Duration (in minutes)"
typeof(string));
table.Columns.Add("Performing HP ID"
typeof(string));
table.Columns.Add("Performing HP ID Issuer"
typeof(string));
table.Columns.Add("Ordering HP ID"
typeof(string));
table.Columns.Add("Ordering HP ID Issuer"
typeof(string));
table.Columns.Add("Referring HP ID"
typeof(string));
table.Columns.Add("Referring HP ID Issuer"
typeof(string));
table.Columns.Add("Supervising HP ID"
typeof(string));
table.Columns.Add("Supervising HP ID Issuer"
typeof(string));
table.Columns.Add("Override Proc. Code"
typeof(string));
table.Columns.Add("Proc. Mod. (1)"
typeof(string));
table.Columns.Add("Proc. Mod. (2)"
typeof(string));
table.Columns.Add("Proc. Mod. (3)"
typeof(string));
table.Columns.Add("Proc. Mod. (4)"
typeof(string));
table.Columns.Add("Override Rev. Code"
typeof(string));
table.Columns.Add("Override Charge Amt."
typeof(string));
table.Columns.Add("Override Service Name"
typeof(string));
table.Columns.Add("Diag. Code (1)"
typeof(string));
table.Columns.Add("Diag. Code (2)"
typeof(string));
table.Columns.Add("Diag. Code (3)"
typeof(string));
table.Columns.Add("Diag. Code (4)"
typeof(string));
table.Columns.Add("Diag. Code (5)"
typeof(string));
table.Columns.Add("Diag. Code (6)"
typeof(string));
table.Columns.Add("Diag. Code (7)"
typeof(string));
table.Columns.Add("Diag. Code (8)"
typeof(string));
table.Columns.Add("Dose Quantity"
typeof(string));
table.Columns.Add("Referral Number"
typeof(string));
table.Columns.Add("Authorization Number"
typeof(string));
table.Columns.Add("Cost"
typeof(string));
table.Columns.Add("ABN Status Date"
typeof(string));
table.Columns.Add("ABN Status Code"
typeof(string));
table.Columns.Add("National Drug Code"
typeof(string));
table.Columns.Add("Item Number"
typeof(string));
table.Columns.Add("Model Number"
typeof(string));
table.Columns.Add("Taxonomy Code"
typeof(string));
table.Columns.Add("Clinic Code"
typeof(string));
table.Columns.Add("Tooth Designation"
typeof(string));
table.Columns.Add("Dent. Srfc. Code1"
typeof(string));
table.Columns.Add("Dent. Srfc. Code2"
typeof(string));
table.Columns.Add("Dent. Srfc. Code3"
typeof(string));
table.Columns.Add("Dent. Srfc. Code4"
typeof(string));
table.Columns.Add("Dent. Srfc. Code5"
typeof(string));
table.Columns.Add("Tooth Status Code"
typeof(string));
table.Columns.Add("Oral Cavity Code1"
typeof(string));
table.Columns.Add("Oral Cavity Code2"
typeof(string));
table.Columns.Add("Oral Cavity Code3"
typeof(string));
table.Columns.Add("Oral Cavity Code4"
typeof(string));
table.Columns.Add("Oral Cavity Code5"
typeof(string));
table.Columns.Add("Placement Status Code"
typeof(string));
table.Columns.Add("Prior Placement Date"
typeof(string));
table.Columns.Add("Podiatry Last PCP Visit Date"
typeof(string));
table.Columns.Add("Hearing And Vision Prescription Date"
typeof(string));
table.Columns.Add("Vision Category Code"
typeof(string));
table.Columns.Add("Vision Certification Condition Indicator"
typeof(string));
table.Columns.Add("Vision Condition Indicator Code1"
typeof(string));
table.Columns.Add("Vision Condition Indicator Code2"
typeof(string));
table.Columns.Add("Vision Condition Indicator Code3"
typeof(string));
table.Columns.Add("Vision Condition Indicator Code4"
typeof(string));
table.Columns.Add("Vision Condition Indicator Code5"
typeof(string));
table.Columns.Add("Dme Certificate of Medical Necessity Transmission Code"
typeof(string));
table.Columns.Add("Dme Certification Type Code"
typeof(string));
table.Columns.Add("Dme Duration(Months)"
typeof(string));
table.Columns.Add("Dme Certification Revision Date"
typeof(string));
table.Columns.Add("Dme initial Certification Date"
typeof(string));
table.Columns.Add("Dme Last Certification Date"
typeof(string));
table.Columns.Add("Dme Length of Medical Necessity Days"
typeof(string));
table.Columns.Add("Dme Rental Price"
typeof(string));
table.Columns.Add("Dme Purchase Price"
typeof(string));
table.Columns.Add("Dme Frequency Code"
typeof(string));
table.Columns.Add("Dme Certification Condition Indicator"
typeof(string));
table.Columns.Add("Dme Condition Indicator Code1"
typeof(string));
table.Columns.Add("Dme Condition Indicator Code2"
typeof(string));
table.Columns.Add("Special Processing Code1"
typeof(string));
table.Columns.Add("Special Processing Code2"
typeof(string));
table.Columns.Add("Special Processing Code3"
typeof(string));
table.Columns.Add("Special Processing Code4"
typeof(string));
table.Columns.Add("Special Processing Code5"
typeof(string));
table.Columns.Add("Special Processing Code6"
typeof(string));
table.Columns.Add("Investigational Device Exempt No"
typeof(string));
table.Columns.Add("Place Of Service Override"
typeof(string));
table.Columns.Add("Procedure Code Description Override"
typeof(string));
table.Columns.Add("Enc. Provider"
typeof(string));
table.Columns.Add("Enc. Location"
typeof(string));
table.Columns.Add("Enc. Matching Strategy Code"
typeof(string));
table.Columns.Add("National Drug Code Quantity"
typeof(string));
table.Columns.Add("RX Number"
typeof(string));
table.Columns.Add("Unit Of Measure Code"
typeof(string));
table.Columns.Add("Auto Charge Rule Text"
typeof(string));
table.Columns.Add("Auto Charge Rule Type Code"
typeof(string));
table.Columns.Add("Service Price Rule Text"
typeof(string));
table.Columns.Add("Error Message"
typeof(string));

}
}
