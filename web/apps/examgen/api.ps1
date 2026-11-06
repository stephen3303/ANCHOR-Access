<#
Exam Generator's web module (read-only first cut), over ExamGen_BE.accdb.
Routes answer under /api/examgen/. Generic parts are the kit core.

Second app of the multi-app proof: its own data file, its own people table
for permission, nothing shared with ANCHOR except who the person is.
#>

# The Access front end's modSignIn: an active tblPeople row is a profile; the
# person's roles come from tblPersonRoles.
Set-KitProfileResolver {
    param($conn, [int]$personId)
    $p = Invoke-KitQuery $conn 'SELECT IsAdmin, IsActive FROM tblPeople WHERE PersonID=?' @($personId)
    if ($p.Count -eq 0 -or -not $p[0].IsActive) { return $null }
    $roles = Invoke-KitQuery $conn ('SELECT r.RoleName FROM tblPersonRoles AS pr INNER JOIN tblRoles AS r ON pr.RoleID=r.RoleID ' +
        'WHERE pr.PersonID=? ORDER BY r.ChainOrder') @($personId)
    $names = @($roles | ForEach-Object { $_.RoleName })
    return @{ isAdmin = [bool]$p[0].IsAdmin; roles = $names; billet = ($names -join ', ') }
}

Set-KitPeopleLister {
    param($conn)
    $rows = Invoke-KitQuery $conn 'SELECT PersonID FROM tblPeople WHERE IsActive ORDER BY PersonID'
    foreach ($row in $rows) { $row['Billet'] = $null; $row['Unit'] = $null }
    return , $rows
}

# What's in the question bank and the exam pipeline, by status.
Add-KitRoute GET '^/overview$' {
    param($r)
    $c = $r.Conn
    $r.Json(200, @{
        questionsByStatus = (Invoke-KitQuery $c ('SELECT s.StatusName AS status, Count(q.QuestionID) AS n FROM tblStatuses AS s ' +
            "LEFT JOIN (SELECT QuestionID, StatusID FROM tblQuestions WHERE IsCurrent) AS q ON s.StatusID=q.StatusID " +
            "WHERE s.AppliesTo='Question' GROUP BY s.StatusID, s.StatusName ORDER BY s.StatusID"))
        examsByStatus     = (Invoke-KitQuery $c ('SELECT s.StatusName AS status, Count(e.ExamID) AS n FROM tblStatuses AS s ' +
            "LEFT JOIN tblExams AS e ON s.StatusID=e.StatusID WHERE s.AppliesTo='Exam' GROUP BY s.StatusID, s.StatusName ORDER BY s.StatusID"))
        topics            = (Invoke-KitQuery $c ('SELECT t.TopicID AS id, t.TopicName AS name, t.DefaultCount AS defaultCount, ' +
            '(SELECT Count(*) FROM tblQuestions AS q WHERE q.TopicID=t.TopicID AND q.IsCurrent) AS questions ' +
            'FROM tblTopics AS t WHERE t.IsActive ORDER BY t.TopicName'))
        templates         = (Invoke-KitQuery $c 'SELECT TemplateID AS id, TemplateName AS name FROM tblTemplates WHERE IsActive ORDER BY TemplateName')
        difficulty        = (Invoke-KitQuery $c 'SELECT LevelName AS name FROM tblDifficultyLevels WHERE IsActive ORDER BY SortOrder')
        classLevels       = (Invoke-KitQuery $c 'SELECT Marking AS marking, PortionMark AS portion FROM tblClassLevels ORDER BY LevelRank')
    })
}
