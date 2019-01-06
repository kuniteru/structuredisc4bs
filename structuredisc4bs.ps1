Param (
    [parameter(mandatory=$true)][ValidateSet("createCSV","applyDisc")][String]$Command,
    [String]$DiscDir = ".",
    [String]$OutDir = ".",
    [parameter(mandatory=$true)][String]$csv,
    [ValidateSet("true", "false")][String]$test = "false",
    [ValidateSet("notUseParser", "UseParser")][String]$csvMode = "notUseParser"
)

if ($test -eq "true") {
    $test = $TRUE
} else {
    $test = $FALSE
}

switch ($Command) {
    "applyDisc" {
        try {
            $DiscDir = Resolve-Path $DiscDir -ErrorAction Stop
        } catch {
            "Folder does not exist. [-DiscDir]"
            exit 1
        }
        try {
            $games = Import-Csv $csv -Encoding Default -ErrorAction Stop
        } catch {
            'File "' + $csv + '" does not exist. [-csv]'
            Write-Debug $_.Exception
            exit 1
        }
        try {
            $asm = [System.Reflection.Assembly]::LoadFile($PSScriptRoot + "\System.Data.sqlite.dll")
        } catch {
            'Could not load "System.Data.sqlite.dll".'
            exit 1
        }
        $db = New-Object System.Data.SQLite.SQLiteConnection
        $db.ConnectionString = "Data Source = ${DiscDir}\System\Databases\regional.db"
        $dbCmd = New-Object System.Data.SQLite.SQLiteCommand
        $dbCmd.Connection = $db
        $db.open()

        $sql = "select max(DISC_ID) as DISC_ID, max(GAME_ID) as GAME_ID from DISC"
        $dbCmd.CommandText = $sql
        $rs = $dbCmd.ExecuteReader()
        if ($rs.Read()) {
            $currentDiscId = $rs["DISC_ID"]
            $currentGameId = $rs["GAME_ID"]
        } else {
            $currentDiscId = 0
            $currentGameId = 0
        }
        $dbCmd.Dispose()

        try {
            $sqlInsertGame = "insert into GAME (GAME_ID, GAME_TITLE_STRING, PUBLISHER_NAME, RELEASE_YEAR, PLAYERS, RATING_IMAGE, GAME_MANUAL_QR_IMAGE, LINK_GAME_ID) " `
                           + "values (@game_id, @game_title_string, @publisher_name, @release_year, @players, @rating_image, @game_manual_qr_image, @link_game_id)"
            $sqlInsertDisc = "insert into DISC (DISC_ID, GAME_ID, DISC_NUMBER, BASENAME) "`
                           + "values (@disc_id, @game_id, @disc_number, @basename)"
            foreach ($game in $games) {
                Try {
                    $gameDir = $OutDir + "\Games\" + [String]($currentGameId + 1)
                    $dummy = New-Item $gameDir -ItemType Directory -ErrorAction Stop
                    $gameDir = (Resolve-Path $gameDir).Path
                } catch {
                    'Folder "' + $gameDir + '" is already exist.'
                    Write-Debug $_.Exception
                    exit 1
                }
                $discs = $game.disc -split ","
                $licFile = $gameDir + "\" + $discs[0] + ".lic"
                $dummy = New-Item $licFile -ItemType File -ErrorAction Stop
                Copy-Item ($PSScriptRoot + "\pcsx.cfg") $gameDir
                Invoke-WebRequest -Uri $game.thumb -OutFile ($gameDir + "\temp")
                $image = [System.Drawing.Image]::FromFile($gameDir + "\temp")
                $image.Save($gameDir + "\" + $discs[0] + ".png", "png")
                $image.Dispose()
                Remove-item ($gameDir + "\temp")

                $dbCmd.CommandText = $sqlInsertGame
                $dbCmd.Parameters.Clear()
                $dummy = $dbCmd.Parameters.AddWithValue("@game_id", $currentGameId + 1)
                $dummy = $dbCmd.Parameters.AddWithValue("@game_title_string", $game.title)
                $dummy = $dbCmd.Parameters.AddWithValue("@publisher_name", $game.maker)
                $dummy = $dbCmd.Parameters.AddWithValue("@release_year", $game.release)
                $dummy = $dbCmd.Parameters.AddWithValue("@players", $game.player)
                $dummy = $dbCmd.Parameters.AddWithValue("@rating_image", "")
                $dummy = $dbCmd.Parameters.AddWithValue("@game_manual_qr_image", "")
                $dummy = $dbCmd.Parameters.AddWithValue("@link_game_id", "")
                $result = $dbCmd.ExecuteNonQuery()

                $dbCmd.CommandText = $sqlInsertDisc
                $dbCmd.Parameters.Clear()
                $discNum = 1
                foreach ($disc in $discs) {
                    try {
                        Move-Item "${DiscDir}\${disc}.cue" "${gameDir}\${disc}.cue" -ErrorAction Stop
                        Move-Item "${DiscDir}\${disc}.bin" "${gameDir}\${disc}.bin" -ErrorAction Stop
                    } catch {
                        "Disc file" + "${DiscDir}\${disc}.*" + "not Exists."
                        Write-Debug $_.Exception
                        exit 1
                    }

                    $dummy = $dbCmd.Parameters.AddWithValue("@disc_id", $currentDiscId + 1)
                    $dummy = $dbCmd.Parameters.AddWithValue("@game_id", $currentGameId + 1)
                    $dummy = $dbCmd.Parameters.AddWithValue("@disc_number", $discNum)
                    $dummy = $dbCmd.Parameters.AddWithValue("@basename", $disc)
                    $currentDiscId ++
                    $discNum ++
                    $result = $dbCmd.ExecuteNonQuery()
                }
                $currentGameId ++
            }
        } catch {
            Write-Debug $_.Exception
        }
    }
    "createCSV" {
        $DatabaseSiteUrl = "https://www.jp.playstation.com/software/title/"
        [array]$games = @()

        $discs = Get-ChildItem $DiscDir -Filter *.cue -File | Sort-Object -Property Name
        :DISCS foreach ($disc in $discs) {
            $game = New-Object PSObject | Select-Object id, disc, title, titleKana, maker, player, release, thumb

            $game.id = $disc.BaseName.ToLower() -replace "[^0-9,^a-z]",""
            $uri = $DatabaseSiteUrl + $game.id + ".html"
            $game.disc = $disc.BaseName

            $response = Invoke-WebRequest -Uri $uri
            try {
                switch ($csvMode) {
                    "notUseParser" {
                        $content = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($response.content)) -creplace "&", "&amp;"
                        # imgエレメントの閉じなし問題への対策
                        $content = $content -replace "<img ([^>]*[^/])>", "<img `$1/>"

                        # brエレメントの閉じなし問題への対策
                        $content = $content -replace "<br>", "<br />"
            
                        # classアトリビュート前にスペースがない問題への対策
                        $content = $content -replace 'id="mdd"class="psc-', 'id="mdd" class="psc-'

                        # 無駄な「\」がある問題への対策
                        $content = $content -replace "<\\/script>", "</script>"

                        $content = [xml]$content
                        $xmlNav = $content.CreateNavigator()
                        $game.title = $xmlNav.Select("//*[@id='softTitle']").Value
                        $game.titleKana = $xmlNav.Select("//*[@id='softTitleKana']").Value
                        $game.maker = $xmlNav.Select("//*[@id='makerName']").Value
                        $game.player = $xmlNav.Select("//*[@id='player']").Value
                        if ($game.player -eq $NULL) {
                            $game.player = "4人"
                        }
                        $game.release = $xmlNav.Select("//*[@id='releaseDate']").Value
                        $game.thumb = $xmlNav.Select("//*[@id='softThumb']").Value

                    }
                    "UseParser" {
                        $title = $response.ParsedHtml.getElementById("softTitle").innerText
                        $game.title = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($title))

                        try {
                            $titleKana = $response.ParsedHtml.getElementById("softTitleKana").innerText
                            $game.titleKana = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($titleKana))
                        } catch {
                            $game.titleKana = ""
                        }
                        try {
                            $maker = $response.ParsedHtml.getElementById("makerName").innerText
                            $game.maker = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($maker))
                        } catch {
                            $game.maker = ""
                        }
                        try {
                            $player = $response.ParsedHtml.getElementById("player").innerText.Trim()
                            $game.player = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($player))
                        } catch {
                            $game.player = "4人"
                        }
                        try {
                            $release = $response.ParsedHtml.getElementById("releaseDate").innerText
                            $game.release = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($release))
                        } catch {
                            $game.release = ""
                        }
                        try {
                            $game.thumb = $response.ParsedHtml.getElementById("softThumb").innerText
                        } catch {
                            Write-Host $_.Exception.Message
                            $game.thumb = ""
                        }
                    }
                }

                $game.maker = $game.maker -replace "\([株有合]\)"
                $game.player = $game.player -match "[0-9]+人$"
                $game.player = $matches.0 -replace "人",""
                $game.release = $game.release.Trim().Substring(0, 4)
                $game.thumb = $game.thumb -replace "\?.*",""
            } catch {
                #Write-Host $_.Exception
            } finally {
                $games += $game

            }

            if ($test -eq $TRUE) {
                break DISCS
            }
        }

        $games = $games | Sort-Object -Property id -Descending
        $games2 = @()
        for ($i = 0; $i -lt $games.Count; $i++) {
            if ($games[$i].title -eq $NULL) {
                try {
                    $currentIdStr = $games[$i].id -match "^[a-z]+"
                    $currentIdStr = $matches.0
                    $currentIdNum = $games[$i].id -match "[0-9]+$"
                    $currentIdNum = [int]$matches.0
                    $nextIdStr = $games[$i + 1].id -match "^[a-z]+"
                    $nextIdStr = $matches.0
                    $nextIdNum = $games[$i + 1].id -match "[0-9]+$"
                    $nextIdNum = [int]$matches.0

                    if (($currentIdStr -ne $nextIdStr) -or ($currentIdNum - 1 -ne $nextIdNum)) {
                        throw "Invalid primary disc Id."
                    } else {
                        $games[$i + 1].disc = $games[$i + 1].disc + "," + $games[$i].disc
                    }

                } catch {
                    "Error!: " + $games[$i].disc + " is sub disc, but it dosn't have main disc."
                }
            } else {
                $games2 += $games[$i]
            }
        }
        $games2 = $games2 | Sort-Object -Property maker,titleKana,title
        $games2 | Export-Csv -Path $csv -Encoding Default -NoTypeInformation
    }
}