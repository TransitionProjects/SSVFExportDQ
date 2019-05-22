__author__ = "David Marienburg"
__maintainer__ = "David Marienburg"
__version__ = "1.1.1"

import pandas as pd
import numpy as np
import zipfile as zf
import datetime

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class ExportDQReport:
    def __init__(self):
        self.export_zip = zf.ZipFile(
            askopenfilename(title="Select the SSVF Export Zip File"),
            "r"
        )

        self.client = pd.read_csv(self.export_zip.open("Client.csv"))
        self.diabilities = pd.read_csv(self.export_zip.open("Disabilities.csv"))
        self.employment_education = pd.read_csv(self.export_zip.open("EmploymentEducation.csv"))
        self.enrollment = pd.read_csv(self.export_zip.open("Enrollment.csv"))
        self.enrollmentcoc = pd.read_csv(self.export_zip.open("EnrollmentCoC.csv"))
        self.exit = pd.read_csv(self.export_zip.open("Exit.csv"))
        self.funder = pd.read_csv(self.export_zip.open("Funder.csv"))
        self.healthanddv = pd.read_csv(self.export_zip.open("HealthAndDV.csv"))
        self.incomebenefits = pd.read_csv(self.export_zip.open("IncomeBenefits.csv"))
        self.project = pd.read_csv(self.export_zip.open("Project.csv"))
        self.projectcoc = pd.read_csv(self.export_zip.open("ProjectCoC.csv"))
        self.services = pd.read_csv(self.export_zip.open("Services.csv"))


    def client_dq(self):
        """
        Translate the self.client sheet into a human readable data quality
        report which explaines in detail what needs to be fixed for a given
        participant.

        A summary report will also be created showing errors counts vs required
        fields organized by the user who created the entry for the participant.

        Both of these repeports will be returned as a tuple of pandas data
        frames.
        """
        # create a local pandas data frame from the self.client object.
        client_df = self.client.copy()

        # create a filled column for ignoring in the np.select process
        client_df["Fill"] = 1

        # create a user readable error for the name data quality column
        name_dq_conditions = [
        	client_df["NameDataQuality"] == 1,
        	client_df["NameDataQuality"] == 99
        ]
        name_dq_choices = [
        	"",
        	"Name Data Quality field is blank"
        ]
        client_df["Name Data Quality Errors"] = np.select(
            name_dq_conditions,
            name_dq_choices
        )

        # create a user readable error for the SSN column
        ssn_conditions = [
        	client_df["SSN"].isna(),
        	(client_df["SSN"] == 0),
        	client_df["SSN"].notna()
        ]
        ssn_choices = [
        	"SSN is blank",
        	"SSN was filled with zeros",
        	""
        ]
        client_df["Social Security Number Errors"] = np.select(
            ssn_conditions,
            ssn_choices
        )

        # create a user readable error for the SSNDataQuality column
        ssn_dq_conditions = [
        	client_df["SSNDataQuality"] == 1,
        	client_df["SSNDataQuality"] == 8,
        	client_df["SSNDataQuality"] == 9,
        	client_df["SSNDataQuality"] == 99
        ]
        ssn_dq_choices = [
        	"",
        	"Client doesn't know was selected when SSN is required",
        	"Client refused is selected for the SSN Data Quality field when SSN is required",
        	"SSN Data Quality field was left blank"
        ]
        client_df["Social Security Number Data Quality Errors"] = np.select(
            ssn_dq_conditions,
            ssn_dq_choices
        )

        # create a user reable error for the DOB column
        dob_conditions = [
            client_df["DOB"].isna(),
            client_df["DOB"].notna()
        ]
        dob_choices = ["Date of Birth field is blank", ""]
        client_df["Date of Birth Errors"] = np.select(
            dob_conditions,
            dob_choices
        )

        # create a user readable error for the DOBDataQuality column
        dob_dq_conditions = [
        	client_df["DOBDataQuality"] == 1,
        	client_df["DOBDataQuality"] == 8,
        	client_df["DOBDataQuality"] == 9,
        	client_df["DOBDataQuality"] == 99
        ]
        dob_dq_choices = [
        	"",
        	"Client doesn't know is selected for the Date of Birth Data Quality field",
        	"Client refused is selected for the Date of Birth Data Quality",
        	"Date of Birth Data Quality field is blank"
        ]
        client_df["Date of Birth Data Quality Errors"] = np.select(
            dob_dq_conditions,
            dob_dq_choices
        )

        # create a user readable error for the race column
        race_conditions = [
        	client_df["RaceNone"].isna(),
        	client_df["RaceNone"] == 9,
        	client_df["RaceNone"] == 99
        ]
        race_choices = [
        	"",
        	"Primary Race is Client Refused",
        	"Primary Race field is blank"
        ]
        client_df["Race Errors"] = np.select(race_conditions, race_choices)

        # create a user readable error for the ethnicity column
        eth_conditions = [
        	client_df["Ethnicity"] == 0,
        	client_df["Ethnicity"] == 1,
        	client_df["Ethnicity"] == 9,
        	client_df["Ethnicity"] == 99
        ]
        eth_choices = [
        	"",
        	"",
        	"Ethnicity field is blank",
        	"Ethnicity field is client refused"
        ]
        client_df["Ethnicity Errors"] = np.select(
            eth_conditions,
            eth_choices
        )

        # create a user readable error for the gender column
        gender_conditions =  [
        	client_df["Gender"] == 0,
        	client_df["Gender"] == 1,
        	client_df["Gender"] == 2,
        	client_df["Gender"] == 3,
        	client_df["Gender"] == 4,
        	client_df["Gender"] == 99
        ]
        gender_choices = [
        	"",
        	"",
        	"",
        	"",
        	"",
        	"Gender field is blank"
        ]
        client_df["Gender Errors"] = np.select(
            gender_conditions,
            gender_choices
        )

        # create a user readable error for the veteran status column
        vet_conditions = [
        	client_df["VeteranStatus"] == 0,
        	client_df["VeteranStatus"] == 1,
        	client_df["VeteranStatus"] == 99
        ]
        vet_choices = [
        	"",
        	"",
        	"U.S. Military Veteran field is blank"
        ]
        client_df["Veteran Status Error"] = np.select(
            vet_conditions,
            vet_choices
        )

        # create a user readable error for the year entered service column
        year_entered_conditions = [
        	(
                (client_df["VeteranStatus"] == 1) &
                client_df["YearEnteredService"].isna()
            ),
        	(
                (client_df["VeteranStatus"] == 1) &
                client_df["YearEnteredService"].notna()
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                client_df["YearEnteredService"].isna()
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                client_df["YearEnteredService"].notna()
            )
        ]
        year_entered_choices = [
        	"Year Entered Service field is blank",
        	"",
            "",
            "Why does this client have a year entered service value when they are not a vet?  This may be an error."
        ]
        client_df["Year Entered Service Error"] = np.select(
            year_entered_conditions,
            year_entered_choices,
            ""
        )

        # create a user readable error for the year entered service column
        year_exited_conditions = [
            (
                (client_df["VeteranStatus"] == 1) &
                client_df["YearSeparated"].isna()
            ),
        	(
                (client_df["VeteranStatus"] == 1) &
                client_df["YearSeparated"].notna()
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                client_df["YearSeparated"].isna()
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                client_df["YearSeparated"].notna()
            )
        ]
        year_exited_choices = [
        	"Year Separated from Service field is blank",
        	"",
            "",
            "Why does this client have a year separated from service value when they are not a vet?  This may be an error."
        ]
        client_df["Year Exited Service Error"] = np.select(
            year_exited_conditions,
            year_exited_choices,
            ""
        )

        # create a user readable error for the military branch column
        branch_conditions = [
            (
                (client_df["VeteranStatus"] == 1) &
                (client_df["MilitaryBranch"].isin([1,2,3,4,5,6]))
            ),
            (
                (client_df["VeteranStatus"] == 1) &
                (client_df["MilitaryBranch"].isna())
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                (client_df["MilitaryBranch"].isin([1,2,3,4,5,6]))
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                (client_df["MilitaryBranch"].isna())
            )
        ]
        branch_choices = [
        	"",
        	"Branch of the Military field is blank",
            "A branch of the military was selected but this client is not a vet.  This may be an error.",
            ""
         ]
        client_df["Military Branch Error"] = np.select(
            branch_conditions,
            branch_choices,
            ""
        )

        # create a user readable error for the discharge status column
        discharge_conditions = [
            (
                (client_df["VeteranStatus"] == 1) &
                (client_df["DischargeStatus"].isin([1,2,3,4,5,6,7]))
            ),
            (
                (client_df["VeteranStatus"] == 1) &
                (client_df["DischargeStatus"].isna())
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                (client_df["DischargeStatus"].isin([1,2,3,4,5,6,7]))
            ),
            (
                (client_df["VeteranStatus"] == 0) &
                (client_df["DischargeStatus"].isna())
            )
        ]
        discharge_choices = [
        	"",
        	"Discharge Status field is blank",
            "A discharge status was selected but this client is not a vet.  This may be an error.",
            ""
        ]
        client_df["Discharge Status Error"] = np.select(
            discharge_conditions,
            discharge_choices,
            ""
        )

        return client_df[[
            "PersonalID",
            "Name Data Quality Errors",
            "Social Security Number Errors",
            "Social Security Number Data Quality Errors",
            "Date of Birth Errors",
            "Date of Birth Data Quality Errors",
            "Race Errors",
            "Ethnicity Errors",
            "Gender Errors",
            "Veteran Status Error",
            "Year Entered Service Error",
            "Year Exited Service Error",
            "Military Branch Error",
            "Discharge Status Error"
        ]].rename(columns={"PersonalID": "Client ID"})


    def education_dq(self):
        """

        """
        # create a local copy of the self.employment_education dataframe
        edu_df = self.employment_education.copy()

        # create an error column for last grade completed
        grade_cond = [
            edu_df["LastGradeCompleted"].isna(),
            (edu_df["LastGradeCompleted"] == 99)
        ]
        grade_choice = [
            "Last Grade Completed field is blank",
            "Last Grade Completed field is set to Data Not Collected"
        ]
        edu_df["Last Grade Completed Error"] =  np.select(
            grade_cond,
            grade_choice,
            ""
        )

        return edu_df[[
            "PersonalID",
            "Last Grade Completed Error"
        ]].rename(columns={"PersonalID": "Client ID"})


    def disabilities_dq(self):
        pass


    def coc_location_dq(self):
        pass


    def enrollment_dq(self):
        """
        Create error report data based on the information contained in the
        enrollment.csv that is part of the zipped SSVF export file.

        The object returned needs to include the CLIENT ID column to make merges
        with the other methods possible.
        """
        # create local copies of the self.enrollment and the self.client data
        # frames
        enroll = self.enrollment.copy()
        client = self.client.copy()

        # create fill columns
        enroll["Fill"] = 1
        client["Fill"] = 1

        # first slice the enroll data frame so that only rows where the value is
        # equal to self head of household then create a new HOHCount column that
        # can be used to identify households with no head or multiple heads
        merged = enroll[
            enroll["RelationshipToHoH"] == 1
        ].groupby(
            by="HouseholdID"
        ).sum().reset_index()[[
            "HouseholdID", "RelationshipToHoH"
        ]].rename(
            index=str,
            columns={"RelationshipToHoH": "HoHCount"}
        ).merge(
            enroll,
            on="HouseholdID",
            how="right"
        )

        # call the numpyselector method to create an error column for the
        # Relationship to HOH data
        hoh_cond = [
        	merged["RelationshipToHoH"].isna(),
        	(merged["RelationshipToHoH"] == 1),
        	(merged["RelationshipToHoH"] > 1)
        ]
        hoh_choice = [
        	"No Head of Household",
        	"",
        	"Multiple Heads of Household"
        ]
        merged["Relationship To Head of Household Error"] = np.select(
            hoh_cond,
            hoh_choice
        )

        # create a merged client and enroll dataframe and use it to identify
        # households with no vet
        vets = client[["PersonalID", "VeteranStatus"]].merge(
            enroll[["PersonalID", "HouseholdID"]],
            on="PersonalID",
            how="right"
        ).groupby(
            by="HouseholdID"
        ).sum().reset_index()[["HouseholdID", "VeteranStatus"]].rename(
            index=str,
            columns={"VeteranStatus": "VetInHH?"}
        )

        # merge this with the merged data frame, then use the numpyselector and
        # np.select
        merged_2 = merged.merge(
            vets,
            on="HouseholdID",
            how="left"
        )
        vets_cond = [
        	(merged_2["VetInHH?"] > 99),
        	((merged_2["VetInHH?"] < 99) & (merged_2["VetInHH?"] > 1)),
        	(merged_2["VetInHH?"] == 1),
        	(merged_2["VetInHH?"] == 0)
        ]
        vets_choices = [
        	"At least one household member's Veteran Status is blank",
        	"",
        	"",
        	"No Veteran in Household"
        ]
        merged_2["Vet In Household Error"] = np.select(vets_cond, vets_choices)

        # create a merged data frame to show vet status and head of household
        # status along with VAMC station number
        vets_2 = client[["PersonalID", "VeteranStatus"]]
        merged_3 = merged_2.merge(
            vets_2,
            on="PersonalID",
            how="left"
        )

        # call the numpyselector method then use np.select to create the VAMC
        # station number error column
        vamc_cond = [
        	(
        		(merged_3["VeteranStatus"] == 1) &
        		(merged_3["RelationshipToHoH"] == 1) &
        		(merged_3["VAMCStation"].isna())
        	),
        	(
        		(merged_3["VeteranStatus"] == 1) &
        		(merged_3["VAMCStation"].isna())
        	),
        	(merged_3["PersonalID"].notna())
        ]
        vamc_choices = [
        	"VAMC Station Number field is blank",
        	"VAMC Station Number is blank but pt is not head of household",
        	""
        ]
        merged_3["VAMC Station Number Error"] = np.select(
            vamc_cond,
            vamc_choices
        )

        # create the non-homeless entry error and the homeless living situation
        # entry error
        if self.project["ProjectName"].str.contains("RRH").all():
            # The self.numpyselector method's dictionaries are incomplete here
            # which is going to be an area that needs maintenance to ensure
            # accuracy
            non_homeless_cond = [
            	(merged_3["LivingSituation"].isna()),
            	(merged_3["LivingSituation"] == 1),
            	(merged_3["LivingSituation"] == 2),
            	(merged_3["LivingSituation"] == 5),
            	(merged_3["LivingSituation"] == 12),
            	(merged_3["LivingSituation"] == 13),
            	(merged_3["LivingSituation"] == 16),
            	(merged_3["LivingSituation"] == 22),
            	(merged_3["LivingSituation"] == 27)
            ]
            non_homeless_choice = [
            	"Residence Prior to Project Entry field was left blank",
            	"",
            	"",
            	"",
            	"",
            	"",
            	"",
            	"Non-homeless living situation at the time of entry into an RRH provider",
            	""
            ]
            merged_3["Living Situation At Entry Error"] = np.select(
                non_homeless_cond,
                non_homeless_choice
            )
        elif self.project["ProjectName"].str.contains("HP").all():
            homeless_cond = [
            	(merged_3["LivingSituation"].isna()),
            	(merged_3["LivingSituation"] == 1),
            	(merged_3["LivingSituation"] == 2),
            	(merged_3["LivingSituation"] == 5),
            	(merged_3["LivingSituation"] == 12),
            	(merged_3["LivingSituation"] == 13),
            	(merged_3["LivingSituation"] == 16),
            	(merged_3["LivingSituation"] == 22),
            	(merged_3["LivingSituation"] == 27)
            ]
            homeless_choice = [
            	"Residence Prior to Project Entry field was left blank",
            	"Homeless living situation at the time of entry into a HP provider",
            	"Homeless living situation at the time of entry into a HP provider",
            	"",
            	"",
            	"",
            	"Homeless living situation at the time of entry into a HP provider",
            	"",
            	""
            ]
            merged_3["Living Situation At Entry"] = np.select(
                homeless_cond,
                homeless_choice
            )
        else:
            merged_3["Living Situation At Entry Error"] = ""

        # call the self.numpyselector method to check for missing values
        # in the disabling condition field
        disable_cond = [
        	merged_3["DisablingCondition"].isna(),
        	(merged_3["DisablingCondition"] == 99),
        	merged_3["DisablingCondition"].notna()
        ]
        disable_choice = [
        	"",
        	"Disabling Condition has been left blank",
        	""
        ]
        merged_3["Disabling Condition Error"] = np.select(
            disable_cond,
            disable_choice
        )

        # create the income as a percent of AMI Error column
        ami_cond = [
        	(merged_3["PercentAMI"].isna() & (merged_3["RelationshipToHoH"] == 1)),
            (merged_3["PercentAMI"].isna() & (merged_3["RelationshipToHoH"] != 1)),
        	merged_3["PercentAMI"].isin([1,2]),
        	(merged_3["PercentAMI"] > 2)
        ]
        ami_choices = [
        	"Participant is the head of household but their % AMI field is blank",
            "",
            "",
        	"AMI is outside of allowable range"
        ]
        merged_3["Income as a Percent AMI Error"] = np.select(
            ami_cond,
            ami_choices
        )

        # create a timeshomeless error column
        times_cond = [
        	(merged_3["TimesHomelessPastThreeYears"] != 99),
        	(merged_3["TimesHomelessPastThreeYears"] == 99),
        	(merged_3["TimesHomelessPastThreeYears"].isna())
        ]
        times_choices = [
        	"",
        	"Times homeless in the last three years field was set to data not collected",
        	""
        ]
        merged_3["Times Homeless in the Last Three Years Error"] = np.select(
            times_cond,
            times_choices
        )

        # create a months homeless error column
        months_cond = [
        	(merged_3["MonthsHomelessPastThreeYears"] == 8),
        	(merged_3["MonthsHomelessPastThreeYears"] == 101),
        	(merged_3["MonthsHomelessPastThreeYears"] == 102),
        	(merged_3["MonthsHomelessPastThreeYears"] == 103),
        	(merged_3["MonthsHomelessPastThreeYears"] == 104),
        	(merged_3["MonthsHomelessPastThreeYears"] == 105),
        	(merged_3["MonthsHomelessPastThreeYears"] == 106),
        	(merged_3["MonthsHomelessPastThreeYears"] == 107),
        	(merged_3["MonthsHomelessPastThreeYears"] == 108),
        	(merged_3["MonthsHomelessPastThreeYears"] == 109),
        	(merged_3["MonthsHomelessPastThreeYears"] == 110),
        	(merged_3["MonthsHomelessPastThreeYears"] == 111),
        	(merged_3["MonthsHomelessPastThreeYears"] == 112),
        	(merged_3["MonthsHomelessPastThreeYears"] == 113),
        	(merged_3["MonthsHomelessPastThreeYears"].isna()),
        	(merged_3["MonthsHomelessPastThreeYears"] == 99)
        ]
        months_choices = [
        	"Client Doesn't Know was selected for the months homeless in the last three years",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"",
        	"Data Not Collected was selected for the months homeless in the last three years"
        ]
        merged_3["Months Homeless in the Last Three Years Error"] = np.select(
            months_cond,
            months_choices
        )

        # create an appoximate date homelessness started error error column
        approx_cond = [
        	(
        		merged_3["DateToStreetESSH"].isna() &
        		(
        			(merged_3["LivingSituation"] == 16) |
        			(merged_3["LivingSituation"] == 27) |
        			(merged_3["LivingSituation"] == 1)
        		)
        	),
        	merged_3["DateToStreetESSH"].notna()
        ]
        approx_choices = [
        	"Client entered from a homeless situation but no approxiamte date homelessness started was entered",
        	""
        ]
        merged_3["Approximate Date Homelessness Started Error"] = np.select(
            approx_cond,
            approx_choices,
            ""
        )

        # create a residence prior to project entry error
        res_cond = [
        	(merged_3["LivingSituation"] == 99),
        	(merged_3["LivingSituation"] != 99)
        ]
        res_choice = [
        	"Data not collected selected in the residence prior to project entry field",
        	""
        ]
        merged_3["Residence Prior to Project Entry Error"] = np.select(
            res_cond,
            res_choice
        )

        return merged_3[[
            "PersonalID",
            "Relationship To Head of Household Error",
            "Vet In Household Error",
            "VAMC Station Number Error",
            "Living Situation At Entry Error",
            "Disabling Condition Error",
            "Income as a Percent AMI Error",
            "Times Homeless in the Last Three Years Error",
            "Months Homeless in the Last Three Years Error",
            "Approximate Date Homelessness Started Error",
            "Residence Prior to Project Entry Error",
            "UserID"
        ]].rename(columns={"PersonalID": "Client ID"})


    def income_benefits_dq(self):
        """

        """
        # make a local copy of the income and benefits dataframe
        income_df = self.incomebenefits.copy().sort_values(
            by=["PersonalID", "DateCreated"],
            ascending=True
        ).drop_duplicates(
            subset="PersonalID",
            keep="first"
        )

        # add the fill column
        income_df["Fill"] = 1

        # create the incomefromanysource error column
        income_1_cond = [
        	(
        		(income_df["IncomeFromAnySource"] == 1) &
        		(income_df["TotalMonthlyIncome"] > 0)
        	),
        	(
        		(income_df["IncomeFromAnySource"] == 1) &
        		(income_df["TotalMonthlyIncome"] == 0)
        	),
        	(income_df["IncomeFromAnySource"] == 0),
        	(income_df["IncomeFromAnySource"].isna()),
        	(income_df["IncomeFromAnySource"] == 99)
        ]
        income_1_choice = [
        	"",
        	"The Income From Any Source Field was set at entry to yes but the HUD Verifications at entry do not show any income",
        	"",
        	"No value was entered into the Income From Any Source Field",
        	"Data not collected was entered into the Income From Any Source Field"
        ]
        income_df["Income From Any Source Error"] = np.select(
            income_1_cond,
            income_1_choice
        )

        # create a benefitsfromanysource error column
        benefits_cond = [
        	(
        		(income_df["BenefitsFromAnySource"] == 1) &
        		((income_df["SNAP"]+income_df["WIC"]+income_df["TANFChildCare"]+income_df["TANFTransportation"]+income_df["OtherTANF"]+income_df["OtherBenefitsSource"]) == 0)
        	),
        	(
        		(income_df["BenefitsFromAnySource"] == 1) &
        		((income_df["SNAP"]+income_df["WIC"]+income_df["TANFChildCare"]+income_df["TANFTransportation"]+income_df["OtherTANF"]+income_df["OtherBenefitsSource"]) > 0)
        	),
        	(
        		(income_df["BenefitsFromAnySource"] == 0) &
        		((income_df["SNAP"]+income_df["WIC"]+income_df["TANFChildCare"]+income_df["TANFTransportation"]+income_df["OtherTANF"]+income_df["OtherBenefitsSource"]) == 0)
        	),
        	(
        		(income_df["BenefitsFromAnySource"] == 0) &
        		((income_df["SNAP"]+income_df["WIC"]+income_df["TANFChildCare"]+income_df["TANFTransportation"]+income_df["OtherTANF"]+income_df["OtherBenefitsSource"]) > 0)
        	)
        ]
        benefits_choice = [
        	"The Benefits From Any Source field is set to yes but none of the HUD Verification fields are set to yes",
        	"",
        	"",
        	"The Benefits From Any Source field is set to no but at least one of the HUD Verification fields is set to yes"
        ]
        income_df["Benefits From Any Source Errors"] = np.select(
            benefits_cond,
            benefits_choice
        )

        # Create an insurace from any source error column
        insur_cond = [
        	(
        	   (income_df["InsuranceFromAnySource"] == 1) &
               ((income_df["Medicaid"]+income_df["Medicare"]+income_df["SCHIP"]+income_df["VAMedicalServices"]+income_df["EmployerProvided"]+income_df["COBRA"]+income_df["PrivatePay"]+income_df["StateHealthIns"]+income_df["IndianHealthServices"]+income_df["OtherInsurance"]) == 0)
        	),
        	(
                (income_df["InsuranceFromAnySource"] == 1) &
                ((income_df["Medicaid"]+income_df["Medicare"]+income_df["SCHIP"]+income_df["VAMedicalServices"]+income_df["EmployerProvided"]+income_df["COBRA"]+income_df["PrivatePay"]+income_df["StateHealthIns"]+income_df["IndianHealthServices"]+income_df["OtherInsurance"]) > 0)
        	),
        	(
                (income_df["InsuranceFromAnySource"] == 0) &
                ((income_df["Medicaid"]+income_df["Medicare"]+income_df["SCHIP"]+income_df["VAMedicalServices"]+income_df["EmployerProvided"]+income_df["COBRA"]+income_df["PrivatePay"]+income_df["StateHealthIns"]+income_df["IndianHealthServices"]+income_df["OtherInsurance"]) == 0)
        	),
        	(
                (income_df["InsuranceFromAnySource"] == 0) &
                ((income_df["Medicaid"]+income_df["Medicare"]+income_df["SCHIP"]+income_df["VAMedicalServices"]+income_df["EmployerProvided"]+income_df["COBRA"]+income_df["PrivatePay"]+income_df["StateHealthIns"]+income_df["IndianHealthServices"]+income_df["OtherInsurance"]) > 0)
        	)
        ]
        insur_choice =[
        	"The Insurance From Any Source field is set to yes but none of the HUD Verification fields are set to yes",
        	"",
        	"",
        	"The Insurance From Any Source field is set to no but at least one of the HUD Verification fields is set to yes"
        ]
        income_df["Insurance From Any Source Error"] = np.select(
            insur_cond,
            insur_choice
        )

        # create a SOAR error Report
        soar_cond = [
            (income_df["ConnectionWithSOAR"].isna()),
            (income_df["ConnectionWithSOAR"] == 8),
            (income_df["ConnectionWithSOAR"] == 9),
            (income_df["ConnectionWithSOAR"] == 99)
        ]
        soar_choices = [
            "The Connection to SOAR field was left blank",
            "The option selected for the Connection to SOAR field is not acceptable to the VA",
            "The option selected for the Connection to SOAR field (Client Refused) is not acceptable to the VA",
            "The option selected for the Connection to SOAR field (Data not collected) is not acceptable to the VA"
        ]
        income_df["Connection to SOAR Error"] = np.select(
            soar_cond,
            soar_choices,
            ""
        )

        return income_df[[
            "PersonalID",
            "Income From Any Source Error",
            "Benefits From Any Source Errors",
            "Insurance From Any Source Error",
            "Connection to SOAR Error"
        ]].rename(columns={"PersonalID": "Client ID"})


    def save_report(self):
        client = self.client_dq()
        enrolled = self.enrollment_dq()
        benefits = self.income_benefits_dq()
        education = self.education_dq()

        merged = client.merge(
            enrolled,
            on="Client ID",
            how="left"
        ).merge(
            benefits,
            on="Client ID",
            how="left"
        ).merge(
            education,
            on="Client ID",
            how="left"
        )

        writer = pd.ExcelWriter(
            asksaveasfilename(
                title="Save the SSVF Export DQ Report",
                defaultextension=".xlsx",
                initialfile="SSVF Export DQ Report {}.xlsx".format(datetime.date.today())
            ),
            engine="xlsxwriter"
        )
        merged[[
            "UserID",
            "Client ID",
            "Name Data Quality Errors",
            "Social Security Number Errors",
            "Social Security Number Data Quality Errors",
            "Date of Birth Errors",
            "Date of Birth Data Quality Errors",
            "Race Errors",
            "Ethnicity Errors",
            "Gender Errors",
            "Relationship To Head of Household Error",
            "Vet In Household Error",
            "Living Situation At Entry Error",
            "Disabling Condition Error",
            "Income as a Percent AMI Error",
            "Times Homeless in the Last Three Years Error",
            "Months Homeless in the Last Three Years Error",
            "Approximate Date Homelessness Started Error",
            "Residence Prior to Project Entry Error",
            "Income From Any Source Error",
            "Benefits From Any Source Errors",
            "Insurance From Any Source Error",
            "Connection to SOAR Error",
            "Last Grade Completed Error",
            "VAMC Station Number Error",
            "Veteran Status Error",
            "Year Entered Service Error",
            "Year Exited Service Error",
            "Military Branch Error",
            "Discharge Status Error"
        ]].sort_values(by="UserID").to_excel(writer, sheet_name="Errors", index=False)
        writer.save()


if __name__ == "__main__":
    a = ExportDQReport()
    a.save_report()
