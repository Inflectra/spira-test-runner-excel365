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

    //enums for association between artifact types we handle in the add-in
    associationEnums: {
        req2req: 1,
        tc2req: 2,
        tc2rel: 3,
        tc2ts: 4
    },

    // enums and various metadata for all artifacts potentially used by the system
    artifacts: [
        // { field: 'requirements', name: 'Requirements', id: 1, hierarchical: true },
        { field: 'testCases', name: 'Test Cases', id: 2, hasFolders: true, hasSubType: true, subTypeId: 7, subTypeName: "TestSteps" },
        // { field: 'incidents', name: 'Incidents', id: 3 },
        // { field: 'releases', name: 'Releases', id: 4, hierarchical: true },
         { field: 'testRuns', name: 'Test Runs', id: 5, disabled: true, hidden: true },
        // { field: 'tasks', name: 'Tasks', id: 6, hasFolders: true },
        //{ field: 'testSteps', name: 'Test Steps', id: 7, disabled: true, hidden: true, isSubType: true },
        { field: 'testSets', name: 'Test Sets', id: 8, hasFolders: true },
        // { field: 'risks', name: 'Risks', id: 14 }
    ],
    //special cases enum
    specialCases: [
        { artifactId: 2, parameter: 'TestStepId', field: 'Description', target: "Call TC:" }
    ]
};

// each artifact has all its standard fields listed, along with important metadata - display name, field type, hard coded values set by system
var templateFields = {
    testRuns: [
        { field: "TestSetId", name: "ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true },
        { field: "Description", name: "Description", type: params.fieldType.text},
        { field: "ReleaseId", name: "Scheduled Release", type: params.fieldType.release },
        {
            field: "TestRunTypeId", name: "Run Type", type: params.fieldType.drop, required: true,
            values: [
                { id: 1, name: "Manual" },
                { id: 2, name: "Automated" }
            ]
        },
        {
            field: "TestSetStatusId", name: "Status", type: params.fieldType.drop, required: true,
            values: [
                { id: 1, name: "Not Started" },
                { id: 2, name: "In Progress" },
                { id: 3, name: "Completed" },
                { id: 4, name: "Blocked" },
                { id: 5, name: "Deferred" }
            ]
        },
        {
            field: "RecurrenceId", name: "Recurrence", type: params.fieldType.drop,
            values: [
                { id: 0, name: "One Time" },
                { id: 1, name: "Hourly" },
                { id: 2, name: "Daily" },
                { id: 3, name: "Weekly" },
                { id: 4, name: "Monthly" },
                { id: 5, name: "Quarterly" },
                { id: 6, name: "Yearly" }
            ]
        },
        { field: "CreatorId", name: "Creator", type: params.fieldType.user },
        { field: "OwnerId", name: "Owner", type: params.fieldType.user },
        {
            field: "TestSetFolderId", name: "Folder", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/test-set-folders",
                idField: "TestSetFolderId",
                nameField: "Name",
                indent: "IndentLevel",
                isProjectBased: true
            }
        },
        { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "PlannedDate", name: "Planned Date", type: params.fieldType.date },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "ExecutionDate", name: "Execution Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "ProjectId", name: "ProjectId", type: params.fieldType.int, isReadOnly: true, isHidden: true },
        { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
    ],

    testCases: [
        { field: "TestCaseId", name: "Case ID", type: params.fieldType.id },
        { field: "TestStepId", name: "Step ID", type: params.fieldType.subId, isSubTypeField: true },
        { field: "Name", name: "Test Case Name", type: params.fieldType.text, required: true, blocksSubType: true },
        { field: "Description", name: "Test Case Description", type: params.fieldType.text, blocksSubType: true },
        { field: "Description", name: "Test Step Description", type: params.fieldType.text, isSubTypeField: true, requiredForSubType: true, extraDataField: "LinkedTestCaseId", extraDataPrefix: "TC" },
        { field: "Position", name: "Position", type: params.fieldType.text, isSubTypeField: true, isReadOnly: true, isHidden: true },
        { field: "ExpectedResult", name: "Test Step Expected Result", type: params.fieldType.text, isSubTypeField: true, requiredForSubType: true },
        { field: "SampleData", name: "Test Step Sample Data", type: params.fieldType.text, isSubTypeField: true },
        {
            field: "TestCasePriorityId", name: "Test Case Priority", type: params.fieldType.drop,
            bespoke: {
                url: "/test-cases/priorities",
                idField: "PriorityId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        {
            field: "TestCaseTypeId", name: "Test Case Type", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/test-cases/types",
                idField: "TestCaseTypeId",
                nameField: "Name",
                isActive: "IsActive"
            }
        },
        {
            field: "TestCaseStatusId", name: "Test Case Status", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/test-cases/statuses",
                idField: "TestCaseStatusId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        { field: "AuthorId", name: "Author", type: params.fieldType.user },
        { field: "OwnerId", name: "Test Case Owner", type: params.fieldType.user },
        { field: "ProjectId", name: "ProjectId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        {
            field: "TestCaseFolderId", name: "Test Case Folder", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/test-folders",
                idField: "TestCaseFolderId",
                nameField: "Name",
                indent: "IndentLevel",
                isProjectBased: true
            }
        },
        { field: "Requirement", name: "New Associated Requirement(s)", type: params.fieldType.text, isAdvanced: true, association: params.associationEnums.tc2req },
        { field: "Release", name: "New Associated Release(s)", type: params.fieldType.text, isAdvanced: true, association: params.associationEnums.tc2rel },
        { field: "TestSet", name: "New Associated Test Set(s)", type: params.fieldType.text, isAdvanced: true, association: params.associationEnums.tc2ts },
        { field: "ComponentIds", name: "Test Case Component", type: params.fieldType.component, isMulti: true },
        { field: "CreationDate", name: "Test Case Creation Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "ConcurrencyDate", name: "Test Case Conc. Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "ConcurrencyDate", name: "Test Step Conc. Date", type: params.fieldType.text, isReadOnly: true, isSubTypeField: true, isHidden: true },
        { field: "ExecutionStatusId", name: "ExecutionStatusId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "IsSuspect", name: "IsSuspect", type: params.fieldType.bool, isReadOnly: true, isHidden: true }
    ],

    // risks: [
    //     { field: "RiskId", name: "ID", type: params.fieldType.id },
    //     { field: "Name", name: "Name", type: params.fieldType.text, required: true },
    //     { field: "Description", name: "Description", type: params.fieldType.text },
    //     { field: "ReleaseId", name: "Release", type: params.fieldType.release },
    //     {
    //         field: "RiskTypeId", name: "Type", type: params.fieldType.drop, required: true,
    //         bespoke: {
    //             url: "/risks/types",
    //             idField: "RiskTypeId",
    //             nameField: "Name",
    //             isActive: "IsActive"
    //         }
    //     },
    //     {
    //         field: "RiskProbabilityId", name: "Probability", type: params.fieldType.drop,
    //         bespoke: {
    //             url: "/risks/probabilities",
    //             idField: "RiskProbabilityId",
    //             nameField: "Name",
    //             isActive: "Active"
    //         }
    //     },
    //     {
    //         field: "RiskImpactId", name: "Impact", type: params.fieldType.drop,
    //         bespoke: {
    //             url: "/risks/impacts",
    //             idField: "RiskImpactId",
    //             nameField: "Name",
    //             isActive: "Active"
    //         }
    //     },
    //     {
    //         field: "RiskStatusId", name: "Status", type: params.fieldType.drop, required: true,
    //         bespoke: {
    //             url: "/risks/statuses",
    //             idField: "RiskStatusId",
    //             nameField: "Name",
    //             isActive: "Active"
    //         }
    //     },
    //     { field: "CreatorId", name: "Creator", type: params.fieldType.user, required: true },
    //     { field: "OwnerId", name: "Owner", type: params.fieldType.user },
    //     { field: "ComponentId", name: "Component", type: params.fieldType.component },
    //     { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true},
    //     { field: "ClosedDate", name: "Closed Date", type: params.fieldType.date },
    //     { field: "ReviewDate", name: "Review Date", type: params.fieldType.date },
    //     { field: "RiskExposure", name: "Risk Exposure", type: params.fieldType.int, isReadOnly: true, isHidden: true },
    //     { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
    //     { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
    // ],
};

function Data() {

    this.user = {
        url: '',
        userName: '',
        api_key: '',
        roleId: 1,
        //TODO this is wrong and should eventually be fixed to limit what user can create or edit client side
        //when add permissions - show in some way to the user what is going on
        // maybe it's as simple as a footnote explaining why projects or artifacts are disabled
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