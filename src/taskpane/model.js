/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 *
 */
export { params, templateFields, tempDataStore, Data };

/*
 *
 * ==========
 * DATA MODEL
 * ==========
 *
 * This holds all the information about the user and all template configuration information. 
 * Future versions of this program can add their artifact to the `templateData` object.
 *
 */

// main store for hard coded enums and metadata that define fields and artifacts
var params = {
    // enums for different artifacts
    artifactEnums: {
        requirements: 1,
        testCases: 2,
        incidents: 3,
        releases: 4,
        testRuns: 5,
        tasks: 6,
        testSteps: 7,
        testSets: 8,
        risks: 14
    },
    //primary field that will be used to create TestCases Shells
    standardShellField: "TestCaseId",
    //secondary field that will be used to create TestRun Shells
    secondaryShellField: "TestSetId",
    //primary key field to associate Incidents and Test Runs
    standardAssociationField: "TestStepId",
    //secondary key field to associate Incidents and Test Runs
    secondaryAssociationField: "TestRunStepId",
    //field that contains the Id of the standard created artifact
    standardResultField: "TestRunId",
    //field that contains the Id of the secondary created artifact
    secondaryResultField: "IncidentId",
    // enums for different types of field - match custom field prop types where relevant
    fieldType: {
        text: 1,
        int: 2,
        num: 3,
        bool: 4,
        date: 5,
        drop: 6,
        multi: 7,
        user: 8,
        // following types don't exist as custom property types as set by Spira - but useful for defining standard field types here
        id: 9,
        subId: 10,
        component: 11, // project level field
        release: 12, // project level field
        arr: 13, // used for comma separated lists in a single cell (eg linked Ids)
        folder: 14 // don't think in reality this will be need
    },
    // enums and various metadata for all artifacts potentially used by the system
    artifacts: [
        {
            field: 'testRuns', name: 'Test Runs', id: 5, conditionField: "IsTestSteps", hasSubType: true, subTypeId: 7,
            subTypeName: "TestSteps", hasSecondaryType: true, SecondaryTypeId: 8, SecondaryTypeField: "TestSetId", secondaryConditionField: "TestRunTypeId",
            secondaryConditionValue: 1, hasSecondaryTarget: true, secondaryTargetId: 2, SecondaryTargetFieldName: "TestCaseId",
            associationField: "TestSetTestCaseId", hasExtraField: true, extraFieldName: "ReleaseId"
        }
    ],
    //extra TC fixed fields (that are not retrieved from Spira, but needed to send to the server)
    extraTcFixedFields: {
        TestRunTypeId: 1,
        StartDate: (function () {
            return new Date(Date.now()).toISOString();
        })(),
        EndDate: (function () {
            var dateOffset = new Date(Date.now()).getTime() + 1 * 60000;
            return new Date(dateOffset).toISOString();
        })(),
        ArtifactTypeId: 5
    },
    extraTsFixedFields: {
        StartDate: (function () {
            return new Date(Date.now()).toISOString();
        })(),
        EndDate: (function () {
            var dateOffset = new Date(Date.now()).getTime() + 1000;
            return new Date(dateOffset).toISOString();
        })(),
    },
};

// each artifact has all its standard fields listed, along with important metadata - display name, field type, hard coded values set by system
var templateFields = {
    testRuns: [
        { field: "TestCaseId", name: "Case ID", type: params.fieldType.id },
        { field: "TestStepId", name: "Step ID", type: params.fieldType.subId, isSubTypeField: true },
        { field: "Name", name: "Test Case Name", type: params.fieldType.text, isReadOnly: true },
        { field: "ReleaseId", name: "Release", type: params.fieldType.release },
        { field: "TestSetId", name: "Set ID", type: params.fieldType.id, shellField: true },
        { field: "TestSetTestCaseId", name: "Set Case Unique ID", type: params.fieldType.id },
        { field: "Description", name: "Test Step Description", type: params.fieldType.text, isSubTypeField: true, extraDataField: "LinkedTestCaseId", extraDataPrefix: "TC", extraIncDesc: true },
        { field: "ExpectedResult", name: "Test Step Expected Result", type: params.fieldType.text, isSubTypeField: true, extraIncDesc: true },
        { field: "SampleData", name: "Test Step Sample Data", type: params.fieldType.text, isSubTypeField: true },
        {
            field: "ExecutionStatusId", name: "Execution Status", type: params.fieldType.drop, isSubTypeField: true, sendField: true,
            values: [
                { id: 1, name: "Failed" },
                { id: 2, name: "Passed" },
                { id: 3, name: "Not Run" },
                { id: 4, name: "Not Applicable" },
                { id: 5, name: "Blocked" },
                { id: 6, name: "Caution" }
            ]
        },
        { field: "ActualResult", name: "Actual Result", type: params.fieldType.text, isSubTypeField: true, sendField: true, extraIncDesc: true },
        { field: "Incident Name", name: "Incident Name", type: params.fieldType.text, isSubTypeField: true, sendField: true, extraArtifact: true },
        { field: "ExecutionStatusId", name: "ExecutionStatusId", type: params.fieldType.text, isReadOnly: true, isHidden: true, },
        { field: "BuildId", name: "BuildId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "EstimatedDuration", name: "EstimatedDuration", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "ActualDuration", name: "ActualDuration", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "ProjectId", name: "ProjectId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "Tags", name: "Tags", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "Position", name: "Position", type: params.fieldType.text, isSubTypeField: true, isReadOnly: true, isHidden: true },
        { field: "ConcurrencyDate", name: "ConcurrencyDate", type: params.fieldType.text, isReadOnly: true, isHidden: true }
    ]
};

function Data() {

    this.user = {
        url: '',
        userName: '',
        api_key: '',
        roleId: 1
    };

    this.projects = [];

    this.currentProject = '';
    this.projectComponents = [];
    this.projectReleases = [];
    this.projectUsers = [];
    this.indentCharacter = ">";

    this.currentArtifact = '';

    this.projectGetRequestsToMake = 3; // users, components, releases
    this.projectGetRequestsMade = 0;

    // counts of artifact specific calls to make
    this.baselineArtifactGetRequests = 1;
    this.artifactGetRequestsToMake = this.baselineArtifactGetRequests;
    this.artifactGetRequestsMade = 0;


    this.artifactData = '';

    this.colors = {
        bgHeader: '#f1a42b',
        bgHeaderSubType: '#fdcb26',
        bgReadOnly: '#eeeeee',
        bgRunField: '#c0fcd6',
        header: '#ffffff',
        headerRequired: '#000000',
        warning: '#ffcccc'
    };

    this.isTemplateLoaded = false;
    this.isGettingDataAttempt = false;
    this.fields = [];
}

function tempDataStore() {
    this.currentProject = '';
    this.projectComponents = [];
    this.projectReleases = [];
    this.projectUsers = [];

    this.currentArtifact = '';
    this.artifactCustomFields = [];
}