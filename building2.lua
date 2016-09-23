-- this file is auto-generated!
-- don't change it manaully.
-- source file: .//example_building.xls
-- created at: Fri Sep 23 11:08:25 2016


local building2 = {}


building2[1] = {
	id = 1,
	name = "house",
	use_money = 1000,
	use_food = 123,
	is_init = "false",
	defense = 100,
	desc = "好",
	cost = {1,1,1},
	whatever = {"name": 333},
}

building2[2] = {
	id = 2,
	name = "church",
	use_money = 5005,
	use_food = 56,
	is_init = "false",
	defense = 513,
	desc = "不好",
	cost = {1,1,1},
	whatever = {"name": 333},
}

building2[3] = {
	id = 3,
	name = "school",
	use_money = 334,
	use_food = 153,
	is_init = "false",
	defense = 100,
	desc = "贵",
	cost = {1,1,100},
	whatever = {"name": 333},
}

building2[4] = {
	id = 4,
	name = "tomb",
	use_money = 11,
	use_food = 153,
	is_init = "false",
	defense = 221,
	desc = "傻瓜才买",
	cost = {1,100,1},
	whatever = {"name": 333},
}

building2.id_by_name= {}
local id_by_name = building2.id_by_name
id_by_name["house"] = 1
id_by_name["church"] = 2
id_by_name["school"] = 3
id_by_name["tomb"] = 4


return building2

