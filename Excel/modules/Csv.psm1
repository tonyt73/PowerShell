<#
 .Synopsis
  CSV helper functions

 .Description
  Functions to:
    * Create consolidated groups
        group name, item, value 1, value 2, value 3
        ,item, value 1, value 2, value 3
        ,item, value 1, value 2, value 3
        Note: the first column is empty. This creates a group.
    * Join columns into groups
        collect columns into a single group(ing)

  .Example
   # TODO: Examples
#>

Set-StrictMode -Version 3.0
Function Get-ConsolidatedGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Groups
    )

    $newData = @()

    # we need to find the new consolidated group name
    # we are assuming the group names are an ordered set ie. date, ordered names etc
    $minGroup = $null
    $maxGroup = $null
    foreach ($key in $Groups.Keys.GetEnumerator()) {
        if (($null -eq $minGroup) -or ($key -lt $minGroup)) {
            $minGroup = $key
        } 
        if (($null -eq $maxGroup) -or ($key -gt $maxGroup)) {
            $maxGroup = $key
        }
    }
    # group name is least ordered to most ordered
    if ($minGroup -ne $maxGroup) {
        # old sprint names are 7 characters (dates) and new sprint names are 3 (major.minor)
        if ($minGroup.Length -ge $maxGroup.Length) {
            $groupName = $minGroup + " - " + $maxGroup
        } else {
            # so, always put the date first
            $groupName = $maxGroup + " - " + $minGroup
        }
    } else {
        $groupName = $minGroup
    }

    # now let's collate the groups into the new consolidated group
    $newGroup = @{}
    # for each key in the groups provided
    foreach ($groupId in $Groups.Keys.GetEnumerator()) {
        # we iterate over that groups data set so we can consolidate the Values for each into 1 group
        foreach ($groupBy in $Groups[$groupId].GetEnumerator()) {
            # if the group by name is not in the new group
            if (-not ($newGroup.ContainsKey($groupBy.Key))) {
                # then lets add it and initialise it with the group by value set
                $newGroup[$groupBy.Key] = $groupBy.Value
            } else {
                # else we take the new value set and accumulate them to the existing group By value set
                for ($i = 2; $i -lt $groupBy.Value.Count; $i++) {
                    $newGroup[$groupBy.Key][$i] += $groupBy.Value[$i]
                }
            }
        }
    }

    # now we consolidate the new groups names in a single group name (ie. no repeats)
    foreach ($kv in $newGroup.GetEnumerator()) {
        $values = $kv.Value
        $values[0] = $groupName
        $newData += $values -Join(",")
        $groupName = ""
    }

    return $newData
}

# This function collates rows into groups
# This first 2 columns must be name types
#   Column 0 - is the group id column
#   Column 1 - is the group by column
# The rest of the columns must be values
Function Join-ColumnsIntoGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Array]$Data,               # you csv data
        [Parameter(Mandatory)]
        [int]$TopGroups,            # the number of groups to remain unaltered (ie. no collation into a single group)
        [Parameter(Mandatory)]
        [int]$CollateGroups,        # the number of groups to collate into a single group
        [switch]$Reverse            # if your data is least significant first, then reverse it for most significant first
    )

    $newData = @($Data[0])
    $dataSize = $Data.Length
    if ($Reverse) {
        # reverse order the csv data
        [Array]::Reverse($Data)
        $dataSize--
    }

    # now consolidate the remaining groups in single groups of m groups
    $tgi = 0    # count the number of top groups processed, before moving to group consolidation
    $groups = [hashtable]@{}
    for ($i = 0; $i -lt $dataSize; $i++) {
        # we take each of the group by items (column 1) and store them in a hash table
        $columns = $Data[$i].Split(',')
        $group = $columns[0]
        $groupBy = $columns[1]
        if (-not ($groups.ContainsKey($group))) {
            # new group
            if ((($tgi -lt $TopGroups) -and ($groups.Count -gt 0)) -or ($groups.Count + 1 -gt $CollateGroups)) {
                $tgi++
                # we've processed all the groups now we collated them and start a new group set
                $newData += Get-ConsolidatedGroup -Groups $groups
                # clear the groups and start again
                $groups = [hashtable]@{}
            }
            $groups[$group] = [hashtable]@{}
        }
        if (-not ($groups[$group].ContainsKey($groupBy))) {
            $groups[$group][$groupBy] = @($columns[0], $columns[1])
            $groups[$group][$groupBy] += @(0) * ($columns.Count - 2)
        } 
        # collate the columns for the group/groupBy
        $g = $groups[$group][$groupBy]
        for ($c = 2; $c -lt $columns.Count; $c++) {
            $g[$c] += [int]$columns[$c]
        }
    }
    $newData += Get-ConsolidatedGroup $groups
    return $newData
}

Function Export-CsvColumns {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Data,
        [Parameter(Mandatory)]
        $Columns
    )

    $newData = @()
    # for each row in the csv data
    $Data.foreach({
        try {
        # get the row as column values
        $values = $_.Split(',')
        # make a new row
        $row = @()
        # for each column index in the columns list
        $Columns.ForEach({
            # add that column value from the column index
            $row += $values[$_]
        })
        # create the new row of data
        $newData += $row -Join(',')
    } catch {
        $_
    }
    }) 

    return $newData
} 

Export-ModuleMember -Function Join-ColumnsIntoGroups
Export-ModuleMember -Function Export-CsvColumns
