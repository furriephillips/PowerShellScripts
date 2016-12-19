<#
.SYNOPSIS
    Periodically shows the status of the Move Request Queue.
.DESCRIPTION
    This script needs to be run in an Exchange Management Shell.

    It counts the Queued/InProgress/Completed/Failed Move Requests,
     pasuses for 60 minutes, then recounts the queues.  You may 
     specify an alternate interval, with " -interval <minutes>".
#>

param (
    [int]$interval = "60"
)

While($true) {
    Clear-Host
    Write-Host "`nIf you wish to see the Move Request Queue Details, you can" -foregroundcolor Yellow
    Write-Host " Ctrl-C out of this script, and use the following commands: -`n" -foregroundcolor Yellow
    $QStates = 'Queued','InProgress','Completed','Failed'
    Foreach ($QState in $QStates) {
        Write-Host Get-MoveRequest -MoveStatus $QState
    }

    Write-Host "`nChecking the Move Request Queue...`n" -foregroundcolor Cyan
    Foreach ($QState in $QStates) {
        Write-Host "$QState  " -nonewline
        (Get-MoveRequest -MoveStatus $QState).count
    }

    $SleepyTime = Write-Output "$interval"
    Write-Host "`nSleeping for $SleepyTime minutes..." -nonewline -foregroundcolor Magenta
    $SleepyMins = 1..$SleepyTime
    $Mins = 0
    Foreach ($Min in $SleepyMins) {
        $Mins++
        Start-Sleep -s 60
        Write-Host "`rSleeping for $SleepyTime minutes... $Mins mins elapsed" -nonewline -foregroundcolor Magenta
    }
}
