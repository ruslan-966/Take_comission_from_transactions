package.cpath = package.cpath .. ";" .. getWorkingFolder() .. "\\lib5" .. _VERSION:sub(_VERSION:len()) .. "\\?.dll"
luacom = require("luacom")
w32 = require("w32")
local s_tab_name = 'trades'
local file_name = '\\files_place\\control_file.txt'						-- Используется для контроля текущей обработанной строки в таблице сделок
local klient_tarifs_file_name = '\\files_place\\KlientTarifs.xlsx'		-- Содержит тарифы, установленные для клиентов
local excel_list_name = 'KlientTarifs'
local correct_file_name = 'U:\\УЦБ_документы\\Quik_deals\\Correct limits\\Lim_correction.lci'	-- Lim_correction_test.lci'
local last_made_line = 0											-- Номер строки в таблице сделок, которая была проанализирована
--Параметры --
local firm_id = 'MC0126900000'
local curr_code = 'SUR'
-- Устанавливаем время выключения программы:
local datetime = os.date("!*t",os.time())
local stop_time = { year = datetime.year,
                   month = datetime.month,
                   day = datetime.day,
                   hour = 18,
                   min = 49,
                   sec = 59
                  }
local seconds_since_epoch_stop_time = os.time(stop_time)
is_run = true

function OnStop()
  is_run = false
end

function create_file(number)
	number = number or 0
	-- Создаем файл в режиме "записи"
	f = io.open(getScriptPath()..file_name,'w')
	-- Закрываем файл
	f:close()
	-- Открываем уже существующий файл в режиме "чтения/записи"
	f = io.open(getScriptPath()..file_name,'r+')
	-- Записываем в файл текущую дату и 0 - количество обработанных строк со сделками
	local curr_date = os.date('%x')
	f:write(curr_date..' '..tostring(number)..'\n')
	f:flush()
	f:close()
	return true
end

function mysplit (inputstr, sep)
	if sep == nil then
		sep = "%s"
    end
    local t={}
    for str in string.gmatch(inputstr, "([^"..sep.."]+)") do
		table.insert(t, str)
    end
    return t
end

function read_all_strings_in_file(file_name)
	file_name = file_name or correct_file_name	--"Correct limits\Lim_correction.lci"
	local arr = {}
	local file = io.open(file_name, "r")
	local str = file:read("l")
	while str do
		table.insert( arr, str)
		str = file:read("l")
	end
	file:close()
	return arr
end

function take_data_from_excel()
	local t = {}
	w32.CoInitialize()
	excel = luacom.CreateObject("Excel.Application")
	excel.Visible = true
	fpath = getScriptPath()..klient_tarifs_file_name
	wb = excel.Workbooks:Open(fpath)
	ws = wb.Worksheets(excel_list_name)
	local i = 2									-- счетчик строк на листе excel
	local val
	repeat
		t[i-1] = {}
		for j = 1, 4 do
			val = excel.Cells(i, j).Value2
			if val == nil then break end 
			t[i-1][j] = val 
		end
		i = i + 1
	until val == nil

	excel.DisplayAlerts = false
	excel:Quit()
	excel = nil
	w32.CoUninitialize()
	return t
end

function round(number)
  if (number - (number % 0.001)) - (number - (number % 0.01)) < 0.005 then
    number = number - (number % 0.01)
  else
    number = (number - (number % 0.01)) + 0.01
  end
 return number
end

function adding_to_lim_file(inputstr, file_name)
	file_name = file_name or correct_file_name
	file = io.open(file_name, "a")
	file:write(inputstr)
	file:close()
	return
end
	
function main()
	local is_time_to_work = true		-- переменные управления временем, логическая
	local count_for_time = 1000			-- значение, при котором проверяется текущее время
	local count_time = 0				-- текущее значение счетчика управления временем
	local klients_tab = take_data_from_excel()
	-- Получаем текущую дату:
	local curr_date = os.date('%x')
	-- Пытаемся открыть файл в режиме "чтения/записи"
	f = io.open(getScriptPath()..file_name,'r+')
	-- Если файл не существует
	if f == nil then 
		create_file()
	else 
		--Читаем первую строку файла, для этого встаем в начало файла
		f:seek("set",0)
		local line = f:read()
		f:close()
		local t = mysplit(line, ' ')	-- Разделяем строку на слова
		if t[1] == curr_date then 		-- Если файл сегодняшний, то
			last_made_line = tonumber(t[2])
		else
			local rez, text = os.remove(getScriptPath()..file_name)	-- Если нет, удаляем его и перезаписывем дату и номер строки
			if rez == nil then										-- Если ошибка при удалении,то выводим
				message('Программа не смогла удалить файл: '..getScriptPath()..file_name..'\n'..'Сообщение: '..text)
				exit()
			end
			create_file()
			last_made_line = 0		
		end
	end
	while is_run and is_time_to_work do
		-- Получаем общее количество строк в таблице сделок
		local number_of_rows = getNumberOf(s_tab_name)
		if number_of_rows > last_made_line then	-- появилась новая строка / строки в таблице
			-- Получаем первую необработанную строку, ее номер соответствует last_made_line
			local corr_trade_string = getItem(s_tab_name, last_made_line)
			-- Если значение получено определяем, установлен ли тариф для этого клиента на этом рынке
			if corr_trade_string ~= nil then
				for i = 1, #klients_tab - 1 do
					str_code = tostring(klients_tab[i][1])
					local sep_numb = str_code:find("[^%d]")
					local new_str_code = str_code:sub(1, sep_numb-1)
					---[[
					if corr_trade_string.class_code == klients_tab[i][2] and
					   corr_trade_string.client_code == new_str_code 
					--]]
					--[[
					if i == 2
					--]]
						then
						--идет блок правки комиссии по клиенту
						local t_limit_sended = read_all_strings_in_file()
						local last_str_from_lim_file = tostring(t_limit_sended[#t_limit_sended])
						local sep_str = mysplit(last_str_from_lim_file, ';')
						local last_entry_number
						for j = #sep_str, 1, -1 do
							local num_input = string.find(sep_str[j], "LIMIT_ID", 1, true)
							if num_input ~= nil then
								last_entry_number = tonumber(string.match (sep_str[j], "%d+"))	-- Получили id последней записи.
								break
							end
						end
						local delta
						if klients_tab[i][4] == '%' then
							delta = - corr_trade_string.value * klients_tab[i][3] / 100.
							delta = round(delta)
						else
							message('Ошибка при расчете банковской комиссии. \nПрограмма остановлена.')
							exit()
						end
						local str_to_send = 'LIMIT_TYPE= MONEY;LIMIT_ID='..tostring(last_entry_number + 1)..';FIRM_ID='.. firm_id
						str_to_send = str_to_send..';TAG=EQTV;CURR_CODE='..curr_code..';CLIENT_CODE='..corr_trade_string.client_code
						str_to_send = str_to_send..';CURRENT_LIMIT='..tostring(delta)..';LIMIT_OPERATION=CORRECT_LIMIT;\n'
						message('Правим комиссию клиенту: \n'..str_to_send)
						adding_to_lim_file(str_to_send)
						break
					end
				end
				last_made_line = last_made_line + 1
				-- Сохраняем новое значение обработанной строки в файл
				local rez, text = os.remove(getScriptPath()..file_name)	-- Если нет, удаляем его и перезаписывем дату и номер строки
				if rez == nil then										-- Если ошибка при удалении,то выводим
					message('Программа не смогла удалить файл: '..getScriptPath()..file_name..'\n'..'Сообщение: '..text)
					exit()
				end
				create_file(last_made_line)
				message('Новый номер обработанной строки: '..tostring(last_made_line)..
						'\nКод клиента: '..corr_trade_string.client_code)
			else
				message("Не удалось получить строку таблицы!")
				exit()
			end  
		end
		sleep(100)
		-- Блок управления временем:
		count_time = count_time + 1
		if count_time >= count_for_time then
			if os.time() >= seconds_since_epoch_stop_time then is_time_to_work = false end
			count_time = 0
		end
	end	
	message('Программа начисления комиссий за сделки прекратила свою работу.')
end 
