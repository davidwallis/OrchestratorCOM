AccessCheck                             void AccessCheck (lHandle As Long, bstrObjectID As String, lRequestedAccess As Long, bstrObjectType As String, pvarGrantedAccess)

AddFolder                               void AddFolder (lHandle As Long, bstrParentID As String, pvarFolderData)

AddIntegrationPack                      void AddIntegrationPack (lHandle As Long, pvarIPDetails)

AddPolicy                               void AddPolicy (lHandle As Long, bstrParentID As String, pvarPolicyData)

AddResource                             void AddResource (lHandle As Long, bstrParentID As String, varObjectData)

AddUserToRuntimeRole                    void AddUserToRuntimeRole (lHandle As Long, userName As String)

ChangeLicense                           void ChangeLicense (bstrProductKey As String, bstrExpirationTime As String)

CheckIn                                 void CheckIn (lHandle As Long, bstrTransactionID As String, bstrObjectID As String, bstrComment As String)

CheckOut                                void CheckOut (lHandle As Long, bstrObjectID As String, bstrOptions As String, pvarObjectData)

ClientConnectSignal                     void ClientConnectSignal (pvarClientDetails)

ConfigureActionServer                   void ConfigureActionServer (lHandle As Long, bstrParentID As String, pvarActionServerDetails)

Connect                                 void Connect (bstrUserName As String, bstrPassword As String, [pvarHandle])

CreatePolicyRequest                     void CreatePolicyRequest (lHandle As Long, bstrPolicyID As String, lFlags As Long, bstrInitializationData As String, bstrTargetActionServer As String, bstrUserName As String, bstrPassword As String, varParentIsWaiting, pvarSeqNumber)

DeleteEvent                             void DeleteEvent (lHandle As Long, bstrUniqueID As String)

DeleteFolder                            void DeleteFolder (lHandle As Long, bstrFolderID As String, lFlags As Long)

DeleteLogEntry                          void DeleteLogEntry (lHandle As Long, bstrPolicyID As String, bstrUniqueID As String)

DeleteObject                            void DeleteObject (lHandle As Long, bstrObjectID As String, bstrUniqueKey As String, lFlags As Long)

DeletePolicy                            void DeletePolicy (lHandle As Long, bstrPolicyID As String, lFlags As Long)

DeletePolicyImages                      void DeletePolicyImages (pvarPolicyIDs)

DeleteResource                          void DeleteResource (lHandle As Long, bstrObjectID As String, bstrResourceID As String)

Disconnect                              void Disconnect (lHandle As Long, bstrUserName As String)

DoesPolicyExist                         void DoesPolicyExist (bstrPolicyID As String)

Find                                    void Find (lHandle As Long, bstrStartLocationID As String, lFlags As Long, bstrSearchString As String, bstrSearchCode As String, pvarFoundResults)

FindPoliciesWithoutImages               void FindPoliciesWithoutImages (lHandle As Long, pvarPolicyIDs)

GetActionServers                        void GetActionServers (lHandle As Long, bstrType As String, pvarActionServerDetails)

GetActionServerTypes                    void GetActionServerTypes (lHandle As Long, pvarActionServerTypes)

GetAuditHistory                         void GetAuditHistory (lHandle As Long, bstrObjectID As String, bstrTransactionID As String, bstrCommand As String, pvarData)

GetCheckOutStatus                       void GetCheckOutStatus (lHandle As Long, bstrObjectID As String, pvarStatus)

GetClientConnections                    void GetClientConnections (lHandle As Long, pvarConnectionDetails)

GetConfigurationIds                     void GetConfigurationIds (lHandle As Long, pvarConfigIds)

GetConfigurationValues                  void GetConfigurationValues (lHandle As Long, bstrConfigID As String, pvarValues)

GetCountersValueAndMarker               void GetCountersValueAndMarker (pvarCounterDetails)

GetCustomStartParameterName             void GetCustomStartParameterName (bstrParamID As String, pvarParamName)

GetCustomStartParameters                void GetCustomStartParameters (pvarParams)

GetDatastoreType                        void GetDatastoreType (pType As Long)

GetEventDetails                         void GetEventDetails (bstrUniqueID As String, pvarEvents)

GetEvents                               void GetEvents (pvarEvents)

GetFolderContents                       void GetFolderContents (lHandle As Long, bstrFolderID As String, pvarFolderContents)

GetFolderPathFromID                     void GetFolderPathFromID (bstrFolderID As String, pvarFolderPath)

GetFolders                              void GetFolders (lHandle As Long, bstrFolderID As String, pvarFolders)

GetInstanceStatusForRequests            void GetInstanceStatusForRequests (lHandle As Long, requestIds, returnRequestIds, returnStatus)

GetIntegrationPacks                     void GetIntegrationPacks (lHandle As Long, pvarIPDetails)

GetLatestPolicyReturnDataDefinition     void GetLatestPolicyReturnDataDefinition (bstrPolicyID As String, pvarVersion, pvarDefinition)

GetLicenseExpirationTime                void GetLicenseExpirationTime (pbstrTime As String)

GetLicenseInformation                   void GetLicenseInformation (lHandle As Long, bstrKey As String, pvarLicenseInformation)

GetLogHistory                           void GetLogHistory (lHandle As Long, bstrPolicyID As String, lFlags As Long, pvarLogData)

GetLogHistoryObjectDetails              void GetLogHistoryObjectDetails (lHandle As Long, bstrInstanceID As String, bstrObjectID As String, bstrInstanceNumber As String, pvarObjectInformation)

GetLogHistoryObjects                    void GetLogHistoryObjects (lHandle As Long, bstrPolicyID As String, bstrInstanceID As String, pvarLogData)

GetLogObjectDetails                     void GetLogObjectDetails (lHandle As Long, bstrObjectID As String, bstrInstanceID As String, bstrInstanceNumber As String, pvarObjectInformation)

GetObjectSecurity                       void GetObjectSecurity (lHandle As Long, bstrObjectID As String, pvarSecurity)

GetObjectTypes                          void GetObjectTypes (lHandle As Long, pvarObjectDetails)

GetPolicyIDFromPath                     void GetPolicyIDFromPath (bstrPolicyPath As String, pvarPolicyID)

GetPolicyInputParameterId               void GetPolicyInputParameterId (lHandle As Long, bstrPolicyID As String, bstrParameterName As String, bstrParameterId As String)

GetPolicyObjectList                     void GetPolicyObjectList (lHandle As Long, bstrPolicyID As String, pvarObjectList)

GetPolicyPathFromID                     void GetPolicyPathFromID (bstrPolicyID As String, pvarPolicyPath)

GetPolicyPublishState                   void GetPolicyPublishState (lHandle As Long, bstrPolicyID As String, plFlags As Long)

GetPolicyRunningState                   void GetPolicyRunningState (varSeqNumber, pvarPolicyRunning)

GetPolicyRunStatus                      void GetPolicyRunStatus (lHandle As Long, bstrExecInstanceID As String, pvarPolicyDetails)

GetProductKey                           void GetProductKey (lHandle As Long, bstrExecInstanceID As String, pvarPolicyDetails)

GetRequestOutputData                    void GetRequestOutputData (lHandle As Long, requestId, returnKeys, returnValues)

GetResources                            void GetResources (lHandle As Long, bstrParentID As String, bstrResourceType As String, pvarObjectData)

GetRunbookTesterPublishedRequests       void GetRunbookTesterPublishedRequests (lHandle As Long, requestIds)

GetVersionInformation                   void GetVersionInformation (pvarVersion)

Initialize                              void Initialize ()

IsPolicyRunning                         void IsPolicyRunning (lHandle As Long, bstrPolicyID As String)

LoadObject                              void LoadObject (lHandle As Long, bstrObjectID As String, pvarObjectData)

LoadPolicy                              void LoadPolicy (lHandle As Long, bstrPolicyID As String, pvarPolicyData)

LoadResource                            void LoadResource (lHandle As Long, bstrResourceID As String, pvarResourceData)

ModifyFolder                            void ModifyFolder (lHandle As Long, bstrFolderID As String, varFolderData)

ModifyObject                            void ModifyObject (lHandle As Long, bstrObjectID As String, bstrUniqueKey As String, varObjectData)

ModifyPolicy                            void ModifyPolicy (lHandle As Long, bstrPolicyID As String, varPolicyData, pvarUniqueKey)

ModifyResource                          void ModifyResource (lHandle As Long, bstrObjectID As String, varObjectData)

MoveObject                              void MoveObject (lHandle As Long, bstrObjectID As String, bstrNewParentID As String, bstrObjectType As String)

PolicyHasMonitor                        void PolicyHasMonitor (lHandle As Long, bstrPolicyID As String, varHasMonitor)

RemoveClientConnection                  void RemoveClientConnection (bstrClientMachine As String)

RemoveIntegrationPack                   void RemoveIntegrationPack (lHandle As Long, pvarIPDetails)

RemoveSatellite                         void RemoveSatellite (bstrName As String)

Replace                                 void Replace(lHandle As Long, bstrStartLocationID As String, lFlags As Long, bstrSearchString As String, bstrSearchCode As String, bstrReplaceString As String, pvarFoundResults)

RetrievePoliciesLinkedToAS              void RetrievePoliciesLinkedToAS (bstrComputer As String, pvarPolicyList)

SetConfigurationValues                  void SetConfigurationValues (lHandle As Long, bstrConfigID As String, varValues)

SetLicenseInformation                   void SetLicenseInformation (lHandle As Long, bstrKey As String, varLicenseInformation)

SetObjectSecurity                       void SetObjectSecurity (lHandle As Long, bstrObjectID As String, bstrSecurity As String)

SetPolicyImage                          void SetPolicyImage (lHandle As Long, bstrPolicyID As String, lImageType As Long, varImageData)

SetPolicyPublishState                   void SetPolicyPublishState (lHandle As Long, bstrPolicyID As String, lFlags As Long)

SetPolicyPublishStateWithParams         void SetPolicyPublishStateWithParams (lHandle As Long, bstrPolicyID As String, lFlags As Long, bstrInitializationData As String)

SetPolicyPublishStateWithParamsAndGetID void SetPolicyPublishStateWithParamsAndGetID (lHandle As Long, bstrPolicyID As String, lFlags As Long, bstrInitializationData As String, pvarSeqNumber)

SetReportingOptions                     void SetReportingOptions (lHandle As Long, bstrOptions As String)

StartSqmNotification                    void StartSqmNotification ()

UndoCheckOut                            void UndoCheckOut (lHandle As Long, bstrObjectID As String, lOptions As Long, pvarObjectData)

UpdateClientActivity                    void UpdateClientActivity (bstrClientServer As String, bstrClientUser As String)

