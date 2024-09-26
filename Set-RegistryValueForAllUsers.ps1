$UserProfilePropertyList = New-Object -TypeName 'System.Collections.Generic.List[System.Object]'
  $UserProfilePropertyList.Add('*')
  $UserProfilePropertyList.Add(@{Name = 'NTAccount'; Expression = {Try {[System.Security.Principal.SecurityIdentifier]::New($_.SID).Translate([System.Security.Principal.NTAccount]).Value} Catch {$Null}}})

$UserProfileList = Get-CIMInstance -Namespace 'Root\CIMv2' -ClassName 'Win32_UserProfile' | Where-Object {($_.Special -eq $False)} | Select-Object -Property ($UserProfilePropertyList)

$UserProfileListCount = ($UserProfileList | Measure-Object).Count

For ($UserProfileListIndex = 0; $UserProfileListIndex -lt $UserProfileListCount; $UserProfileListIndex++)
  {
      $UserProfile = $UserProfileList[$UserProfileListIndex]

      $RegistryTableList = New-Object -TypeName 'System.Collections.Generic.List[System.Collections.Specialized.OrderedDictionary]'
        $RegistryTableList.Add([Ordered]@{Enabled = $True; KeyPath = "HKEY_USERS\$($UserProfile.SID)\Software\Microsoft\Office\16.0\Outlook\Options\General"; ValueName = 'HideNewOutlookToggle'; DesiredValue = '1'; ValueKind = [Microsoft.Win32.RegistryValueKind]::DWord})
        #$RegistryTableList.Add([Ordered]@{Enabled = $True; KeyPath = "HKEY_USERS\$($UserProfile.SID)\Software\Microsoft\Office\16.0\Outlook\Options\General"; ValueName = 'HideNewOutlookToggle'; DesiredValue = '1'; ValueKind = [Microsoft.Win32.RegistryValueKind]::DWord})

      $RegistryTableListCount = $RegistryTableList.Count
                                          
      $RegistryTableListCounter = 1
                                          
      For ($RegistryTableListIndex = 0; $RegistryTableListIndex -lt $RegistryTableListCount; $RegistryTableListIndex++)
        {
            Try
              {
                  $RegistryTable = $RegistryTableList[$RegistryTableListIndex]

                  Switch ($RegistryTable.Enabled)
                    {
                        {($_ -eq $True)}
                          {
                              $RegistryValue = [Microsoft.Win32.Registry]::GetValue($RegistryTable.KeyPath, $RegistryTable.ValueName, 'ValueDoesNotExist')
                              
                              Switch (($RegistryValue -ieq 'ValueDoesNotExist') -or ($RegistryValue -ine $RegistryTable.DesiredValue))
                                {                                    
                                    {($_ -eq $True)}
                                      {                                          
                                          $LogMessage = "Attempting to perform registry configuration $($RegistryTableListCounter) of $($RegistryTableListCount). Please Wait..."
                                          Write-Verbose -Message ($LogMessage) -Verbose

                                          ForEach ($RegistryTableKVP In $RegistryTable.GetEnumerator())
                                            {
                                                Switch ($RegistryTableKVP.Key)
                                                  {
                                                      {($_ -inotin @('Enabled'))}
                                                        {
                                                            $LogMessage = "$($RegistryTableKVP.Key): $($RegistryTableKVP.Value)"
                                                            Write-Verbose -Message ($LogMessage) -Verbose
                                                        }
                                                  }
                                            }

                                          $Null = [Microsoft.Win32.Registry]::SetValue($RegistryTable.KeyPath, $RegistryTable.ValueName, $RegistryTable.DesiredValue, $RegistryTable.ValueKind)
                                      }
                                }
                          }
                    }
              }
            Catch
              {
                  $ExceptionPropertyDictionary = New-Object -TypeName 'System.Collections.Specialized.OrderedDictionary'
                    $ExceptionPropertyDictionary.Message = $_.Exception.Message
                    $ExceptionPropertyDictionary.Category = $_.Exception.ErrorRecord.FullyQualifiedErrorID
                    $ExceptionPropertyDictionary.Script = Try {[System.IO.Path]::GetFileName($_.InvocationInfo.ScriptName)} Catch {$Null}
                    $ExceptionPropertyDictionary.LineNumber = $_.InvocationInfo.ScriptLineNumber
                    $ExceptionPropertyDictionary.LinePosition = $_.InvocationInfo.OffsetInLine
                    $ExceptionPropertyDictionary.Code = $_.InvocationInfo.Line.Trim()

                  ForEach ($ExceptionProperty In $ExceptionPropertyDictionary.GetEnumerator())
                    {
                        $LogMessage = "$($ExceptionProperty.Key): $($ExceptionProperty.Value)"
                        Write-Warning -Message ($LogMessage) -Verbose 
                    }
              }
            Finally
              {
                  $RegistryTableListCounter++
              }
        }
  }
