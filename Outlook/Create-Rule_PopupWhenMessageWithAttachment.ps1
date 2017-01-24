$outlook = New-Object -ComObject outlook.application
$rules = $outlook.Session.DefaultStore.GetRules()
$olRuleType = "Microsoft.Office.Interop.Outlook.OlRuleType" -as [type]
$rule = $rules.Create("Popup Alert - Message with attanchment",$olRuleType::OlRuleReceive)
$FromCondition = $rule.Conditions.HasAttachment
$FromCondition.Enabled = $true
$RuleAction = $rule.Actions.NewItemAlert
$RuleAction.Text = "You received "
$RuleAction.Enabled = $true
$rules.Save()
