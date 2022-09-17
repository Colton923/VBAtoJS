Global;
let SteelMode: boolean;
let epsilon: number = 0.0001;
let offset_constant: number = 21;
// data declarations
PrivateType;
item_type_data;
let id: number;
let Width: number;
let Height: number;
let Area: number;
let rotatable: boolean;
let mandatory: number;
let profit: number;
let number_requested: number;
let sort_criterion: number;
let Placement: string;
let Size: string;
EndType;
PrivateType;
item_list_data;
let num_item_types: number;
let total_number_of_items: number;
let item_types: item_type_data[];
EndType;
let item_list: item_list_data;
PrivateType;
bin_type_data;
let type_id: number;
let bin_name: string;
let Width: number;
let Height: number;
let Area: number;
let mandatory: number;
let cost: number;
let number_available: number;
let Placement: string;
let Size: string;
EndType;
PrivateType;
bin_list_data;
let num_bin_types: number;
let bin_types: bin_type_data[];
EndType;
let bin_list: bin_list_data;
PrivateType;
compatibility_data;
let item_to_item: boolean[];
let bin_to_item: boolean[];
EndType;
let compatibility_list: compatibility_data;
PrivateType;
item_location;
let sw_x: number;
let sw_y: number;
let max_x: number;
let max_y: number;
//  for Guillotine cuts
EndType;
PrivateType;
item_in_bin;
let item_type: number;
let item_name: string;
let rotated: boolean;
let mandatory: number;
let sw_x: number;
let sw_y: number;
let ne_x: number;
let ne_y: number;
let max_x: number;
let max_y: number;
let first_cut_direction: number;
let cut_length: number;
let Placement: string;
let Size: string;
EndType;
PrivateType;
bin_data;
let type_id: number;
let bin_name: string;
let Width: number;
let Height: number;
let Area: number;
let cost: number;
let item_cnt: number;
let mandatory: number;
let items: item_in_bin[];
let addition_points: item_location[];
let repack_item_count: number[];
let area_packed: number;
EndType;
PrivateType;
solution_data;
let num_bins: number;
let feasible: boolean;
let net_profit: number;
let total_area: number;
let total_distance: number;
let total_area_utilization: number;
let item_type_order: number[];
let rotation_order: number[];
let first_cut_direction: number[];
let bin: bin_data[];
let unpacked_item_count: number[];
EndType;
PrivateType;
instance_data;
let item_item_compatibility_worksheet: boolean;
let bin_item_compatibility_worksheet: boolean;
let guillotine_cuts: boolean;
let global_upper_bound: number;
// '' max profit obtainable
EndType;
let instance: instance_data;
PrivateType;
solver_option_data;
let CPU_time_limit: number;
let item_sort_criterion: number;
let show_progress: boolean;
EndType;
let solver_options: solver_option_data;


    private SortBins(solution: solution_data) {
        let i: number;
        let j: number;
        let candidate_index: number;
        let max_mandatory: number;
        let max_area_packed: number;
        let min_ratio: number;
        let swap_bin: bin_data;
        // insertion sort
        if ((Rnd < 0.8)) {
            // insertion sort
            // With...
            for (i = 1; (i <= solution.num_bins); i++) {
                candidate_index = i;
                max_mandatory = solution.bin;
                i.mandatory;
                max_area_packed = solution.bin;
                i.area_packed;
                min_ratio = solution.bin;
                (i.cost / solution.bin);
                i.Area;
                for (j = (i + 1); (j <= solution.num_bins); j++) {
                    if ((solution.bin[j].mandatory > max_mandatory)) {
                        ((solution.bin[j].mandatory == max_mandatory)
                                    & (solution.bin[j].area_packed
                                    > (max_area_packed + epsilon)));
                        ((solution.bin[j].mandatory == 0)
                                    & ((max_mandatory == 0)
                                    & (solution.bin[j].area_packed
                                    > (max_area_packed - epsilon))));
                        (solution.bin[j].cost / solution.bin)[j].Area;
                        min_ratio;
                        candidate_index = j;
                        max_mandatory = solution.bin;
                        j.mandatory;
                        max_area_packed = solution.bin;
                        j.area_packed;
                        min_ratio = solution.bin;
                        (j.cost / solution.bin);
                        j.Area;
                    }

                }

                if ((candidate_index != i)) {
                    swap_bin = solution.bin;
                    candidate_index.bin(candidate_index) = solution.bin;
                    i.bin(i) = swap_bin;
                }

            }

        }
        else {
            // With...
            for (i = 1; (i <= solution.num_bins); i++) {
                candidate_index = Int(((((solution.num_bins - i)
                                + 1)
                                * Rnd)
                                + i));
                if ((candidate_index != i)) {
                    swap_bin = solution.bin;
                    candidate_index.bin(candidate_index) = solution.bin;
                    i.bin(i) = swap_bin;
                }

            }

        }

    }

    private PerturbSolution(solution: solution_data) {
        let i: number;
        let j: number;
        let k: number;
        let swap_long: number;
        let bin_emptying_probability: number;
        let item_removal_probability: number;
        let repack_flag: boolean;
        let continue_flag: boolean;
        let empty_type_probability: number;
        let empty_type: number;
        empty_type_probability = Rnd;
        if ((empty_type_probability < 0.5)) {
            empty_type = 0;
        }
        else {
            empty_type = 1;
        }

        for (i = 1; (i <= solution.num_bins); i++) {
            // With...
            if ((empty_type == 0)) {
                bin_emptying_probability = (1 - (0.8
                            * (solution.bin(i).area_packed / solution.bin(i).Area)));
                item_removal_probability = 0;
            }
            else if ((empty_type == 1)) {
                bin_emptying_probability = 0.2;
                item_removal_probability = (0.1
                            + (Rnd * 0.1));
            }

            if ((solution.bin(i).item_cnt > 0)) {
                if ((Rnd() < bin_emptying_probability)) {
                    // empty the bin
                    for (j = 1; (j <= solution.bin(i).item_cnt); j++) {
                        solution.unpacked_item_count(solution.bin(i).items, j.item_type) = (solution.unpacked_item_count(solution.bin(i).items, j.item_type) + 1);
                        solution.net_profit = (solution.net_profit - item_list.item_types(solution.bin(i).items, j.item_type).profit);
                    }

                    solution.net_profit = (solution.net_profit + solution.bin(i).cost);
                    0.addition_points(1).max_x = solution.bin(i).Width.addition_points;
                    0.addition_points(1).sw_y = solution.bin(i).Width.addition_points;
                    0.addition_points(1).sw_x = solution.bin(i).Width.addition_points;
                    0.area_packed = solution.bin(i).Width.addition_points;
                    (solution.total_area - solution.bin(i).area_packed.item_cnt) = solution.bin(i).Width.addition_points;
                    solution.total_area = solution.bin(i).Width.addition_points;
                    1.max_y = solution.bin(i).Height;
                }
                else {
                    repack_flag = false;
                    for (j = 1; (j <= solution.bin(i).item_cnt); j++) {
                        if (((solution.feasible == false)
                                    && (solution.bin(i).items[j].mandatory == 0))) {
                            (Rnd() < item_removal_probability);
                            solution.unpacked_item_count(solution.bin(i).items, j.item_type) = (solution.unpacked_item_count(solution.bin(i).items, j.item_type) + 1);
                            (solution.net_profit - item_list.item_types(solution.bin(i).items, j.item_type).profit.items(j).item_type) = 0;
                            solution.net_profit = 0;
                            repack_flag = true;
                        }

                    }

                    if ((repack_flag == true)) {
                        for (j = 1; (j <= solution.bin(i).item_cnt); j++) {
                            if (solution.bin(i).items) {
                                (j.item_type > 0);
                                solution.net_profit = (solution.net_profit - item_list.item_types(solution.bin(i).items, j.item_type).profit);
                            }

                        }

                        solution.net_profit = (solution.net_profit + solution.bin(i).cost);
                        solution.total_area = (solution.total_area - solution.bin(i).area_packed);
                        for (j = 1; (j <= 0); j++) {
                        }

                        for (j = 1; (j <= solution.bin(i).item_cnt); j++) {
                            if (solution.bin(i).items) {
                                (j.item_type > 0);
                                solution.bin(i).repack_item_count;
                                solution.bin(i).items[j].item_type;
                                solution.bin(i).repack_item_count;
                                solution.bin(i).items[j].item_type;
                                1;
                            }

                        }

                        0.addition_points(1).max_x = solution.bin(i).Width.addition_points;
                        0.addition_points(1).sw_y = solution.bin(i).Width.addition_points;
                        0.addition_points(1).sw_x = solution.bin(i).Width.addition_points;
                        0.item_cnt = solution.bin(i).Width.addition_points;
                        j.area_packed = solution.bin(i).Width.addition_points;
                        1.max_y = solution.bin(i).Height;
                        // repack now
                        for (j = 1; (j <= item_list.num_item_types); j++) {
                            continue_flag = true;
                            while ((solution.bin(i).repack_item_count[j] > 0)) {
                                (continue_flag == true);
                                continue_flag = AddItemToBin(solution, i, j, 2);
                            }

                            //  put the remaining items in the unpacked items list
                            solution.unpacked_item_count(j) = (solution.unpacked_item_count(j) + solution.bin(i).repack_item_count);
                            j.repack_item_count(j) = 0;
                        }

                    }

                }

            }

        }

        // change the preferred rotation order randomly
        for (i = 1; (i <= item_list.num_item_types); i++) {
            if ((Rnd() < 0.2)) {
                if ((solution.rotation_order(i, 1) == 1)) {
                    solution.rotation_order(i, 1) = 0;
                    solution.rotation_order(i, 2) = 1;
                }
                else {
                    solution.rotation_order(i, 1) = 1;
                    solution.rotation_order(i, 2) = 0;
                }

            }

        }

        // change the first cut direction randomly
        for (i = 1; (i <= item_list.num_item_types); i++) {
            if ((Rnd() < 0.2)) {
                if ((solution.first_cut_direction(i) == 1)) {
                    solution.first_cut_direction(i) = 0;
                }
                else {
                    solution.first_cut_direction(i) = 1;
                }

            }

        }

        // change the item order randomly
        for (i = 1; (i <= item_list.num_item_types); i++) {
            if ((Rnd < 0.1)) {
                j = Int(((((item_list.num_item_types - i)
                                + 1)
                                * Rnd)
                                + i));
                //  the order to swap with
                swap_long = solution.item_type_order(i);
                solution.item_type_order(i) = solution.item_type_order(j);
                solution.item_type_order(j) = swap_long;
            }

        }

    }

    private AddItemToBin(solution: solution_data, bin_index: number, item_type_index: number, add_type: number) {
        let i: number;
        let j: number;
        let rotation: number;
        let sw_x: number;
        let sw_y: number;
        let ne_x: number;
        let ne_y: number;
        let min_x: number;
        let min_y: number;
        let candidate_position: number;
        let candidate_rotation: number;
        // With...
        min_x = (solution.bin(bin_index).Width + 1);
        min_y = (solution.bin(bin_index).Height + 1);
        candidate_position = 0;
        // area size check
        if ((solution.bin(bin_index).area_packed
                    + (item_list.item_types(item_type_index).Area > solution.bin(bin_index).Area))) {
            /* Warning! GOTO is not Implemented */}

        // item to item compatibility check
        for (rotation = 1; (rotation <= 2); rotation++) {
            if (((solution.rotation_order(item_type_index, rotation) == 1)
                        && (item_list.item_types(item_type_index).rotatable == false))) {
                /* Warning! GOTO is not Implemented */}

            for (i = 1; (i
                        <= (solution.bin(bin_index).item_cnt + 1)); i++) {
                sw_x = solution.bin(bin_index).addition_points;
                i.sw_x;
                sw_y = solution.bin(bin_index).addition_points;
                i.sw_y;
                if ((solution.rotation_order(item_type_index, rotation) == 0)) {
                    ne_x = (sw_x + item_list.item_types(item_type_index).Width);
                    ne_y = (sw_y + item_list.item_types(item_type_index).Height);
                }
                else {
                    ne_x = (sw_x + item_list.item_types(item_type_index).Height);
                    ne_y = (sw_y + item_list.item_types(item_type_index).Width);
                }

                // check the feasibility of all four corners, w.r.t to the other items
                if (((ne_x
                            > (solution.bin(bin_index).Width + epsilon))
                            || (ne_y
                            > (solution.bin(bin_index).Height + epsilon)))) {
                    /* Warning! GOTO is not Implemented */}

                if ((instance.guillotine_cuts == true)) {
                    if (((ne_x > solution.bin(bin_index).addition_points)[i].max_x + epsilon)) {
                        ((ne_y > solution.bin(bin_index).addition_points)[i].max_y + epsilon);
                        /* Warning! GOTO is not Implemented */}

                    for (j = 1; (j <= solution.bin(bin_index).item_cnt); j++) {
                        if (((sw_x < solution.bin(bin_index).items)[j].ne_x - epsilon)) {
                            ((ne_x > solution.bin(bin_index).items)[j].sw_x + epsilon);
                            ((ne_y > solution.bin(bin_index).items)[j].sw_y + epsilon);
                            ((sw_y < solution.bin(bin_index).items)[j].ne_y - epsilon);
                            /* Warning! GOTO is not Implemented */j;
                            // no conflicts at this point
                            if (((sw_y < min_y)
                                        || ((sw_y
                                        <= (min_y + epsilon))
                                        && (sw_x < min_x)))) {
                                min_x = sw_x;
                                min_y = sw_y;
                                candidate_position = i;
                                candidate_rotation = solution.rotation_order(item_type_index, rotation);
                            }

                            /* Warning! Labeled Statements are not Implemented */i;
                            /* Warning! Labeled Statements are not Implemented */rotation;
                        }

                        /* Warning! Labeled Statements are not Implemented */if ((candidate_position == 0)) {
                            AddItemToBin = false;
                        }
                        else {
                            // With...
                            addition_points(candidate_position).sw_y;
                            (solution.bin(bin_index).item_cnt + 1.items(solution.bin(bin_index).item_cnt).item_type) = solution.bin(bin_index).addition_points;
                            solution.bin(bin_index).item_cnt = solution.bin(bin_index).addition_points;
                            if ((candidate_rotation == 1)) {
                                items(., item_cnt).rotated = true;
                            }
                            else {
                                items(., item_cnt).rotated = false;
                            }

                            item_list.item_types(item_type_index).Placement.items(., item_cnt).Size = item_list.item_types(item_type_index).Size;
                            item_list.item_types(item_type_index).mandatory.items(., item_cnt).Placement = item_list.item_types(item_type_index).Size;
                            items(., item_cnt).mandatory = item_list.item_types(item_type_index).Size;
                            if ((candidate_rotation == 0)) {
                                (items(., item_cnt).sw_y + item_list.item_types(item_type_index).Height);
                            }
                            else {
                                (items(., item_cnt).sw_y + item_list.item_types(item_type_index).Width);
                            }

                            (area_packed + item_list.item_types(item_type_index).Area);
                            if ((instance.guillotine_cuts == true)) {
                                items(., item_cnt).first_cut_direction = solution.first_cut_direction(item_type_index);
                                items(., item_cnt).first_cut_direction = 0;
                                addition_points(candidate_position).max_y;
                                items(., item_cnt).first_cut_direction = 1;
                                addition_points(candidate_position).max_x;
                                items(., item_cnt).cut_length = 0;
                            }
                            else {
                                items(., item_cnt).cut_length = 0;
                                items(., item_cnt).first_cut_direction = 0;
                                (addition_points(candidate_position).max_y - epsilon);
                                items(., item_cnt).sw_x;
                                // .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                            }

                            (addition_points(candidate_position).max_x - epsilon);
                            items(., item_cnt).sw_y;
                            // .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                        }

                        if (solution.bin(bin_index).items) {
                            (solution.bin(bin_index).item_cnt.ne_x < solution.bin(bin_index).addition_points);
                            (candidate_position.max_x - epsilon);
                            solution.bin(bin_index).items;
                            solution.bin(bin_index).item_cnt.cut_length = solution.bin(bin_index).items;
                            (solution.bin(bin_index).item_cnt.cut_length
                                        + (solution.bin(bin_index).addition_points[candidate_position].max_y - solution.bin(bin_index).items));
                            solution.bin(bin_index).item_cnt.sw_y;
                            // .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                        }

                        if (solution.bin(bin_index).items) {
                            (solution.bin(bin_index).item_cnt.ne_y < solution.bin(bin_index).addition_points);
                            (candidate_position.max_y - epsilon);
                            solution.bin(bin_index).items;
                            solution.bin(bin_index).item_cnt.cut_length = solution.bin(bin_index).items;
                            (solution.bin(bin_index).item_cnt.cut_length
                                        + (solution.bin(bin_index).items[solution.bin(bin_index).item_cnt].ne_x - solution.bin(bin_index).items));
                            solution.bin(bin_index).item_cnt.sw_x;
                            // .items(.item_cnt).cut_length = .items(.item_cnt).cut_length + 1
                        }

                        if (solution.bin(bin_index).items) {
                            solution.bin(bin_index).item_cnt.max_x = solution.bin(bin_index).addition_points;
                            candidate_position.max_x.items(solution.bin(bin_index).item_cnt).max_y = solution.bin(bin_index).addition_points;
                            candidate_position.max_y;
                        }

                        if ((add_type == 2)) {
                            solution.bin(bin_index).repack_item_count;
                            item_type_index = solution.bin(bin_index).repack_item_count;
                            (item_type_index - 1);
                        }

                        // update the addition points
                        for (i = candidate_position; (i <= solution.bin(bin_index).addition_points); i++) {
                            (i + 1);
                        }

                        i.addition_points(solution.bin(bin_index).item_cnt).sw_x = solution.bin(bin_index).items;
                        solution.bin(bin_index).item_cnt.ne_x.addition_points(solution.bin(bin_index).item_cnt).sw_y = solution.bin(bin_index).items;
                        solution.bin(bin_index).item_cnt.sw_y.addition_points((solution.bin(bin_index).item_cnt + 1)).sw_x = solution.bin(bin_index).items;
                        solution.bin(bin_index).item_cnt.sw_x.addition_points((solution.bin(bin_index).item_cnt + 1)).sw_y = solution.bin(bin_index).items;
                        solution.bin(bin_index).item_cnt.ne_y;
                        if ((instance.guillotine_cuts == true)) {
                            if (solution.bin(bin_index).items) {
                                solution.bin(bin_index).item_cnt.first_cut_direction = 0;
                                solution.bin(bin_index).addition_points;
                                solution.bin(bin_index).item_cnt.max_x = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.max_x.addition_points(solution.bin(bin_index).item_cnt).max_y = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.ne_y.addition_points((solution.bin(bin_index).item_cnt + 1)).max_x = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.max_x.addition_points((solution.bin(bin_index).item_cnt + 1)).max_y = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.max_y;
                            }
                            else {
                                solution.bin(bin_index).addition_points;
                                solution.bin(bin_index).item_cnt.max_x = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.max_x.addition_points(solution.bin(bin_index).item_cnt).max_y = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.max_y.addition_points((solution.bin(bin_index).item_cnt + 1)).max_x = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.ne_x.addition_points((solution.bin(bin_index).item_cnt + 1)).max_y = solution.bin(bin_index).items;
                                solution.bin(bin_index).item_cnt.max_y;
                            }

                        }

                        // With...
                        // With...
                        // update the profit
                        if (solution.bin) {
                            bin_index.item_cnt = 1;
                            solution.net_profit = (solution.net_profit
                                        + (item_list.item_types(item_type_index).profit - solution.bin));
                            bin_index.cost;
                        }
                        else {
                            solution.net_profit = (solution.net_profit + item_list.item_types(item_type_index).profit);
                        }

                        // update the area per bin and the total area
                        solution.total_area = (solution.total_area + item_list.item_types(item_type_index).Area);
                        // update the unpacked items
                        if ((add_type == 1)) {
                            solution.unpacked_item_count;
                            item_type_index = solution.unpacked_item_count;
                            (item_type_index - 1);
                        }

                        AddItemToBin = true;
                        // ''''''''''''''''''''' Sub for reading data from the item sheet, writing to item type
                        GetItemData((<Collection>(InputCollection)), Optional, (<boolean>(SteelClass)));
                        item_list.num_item_types = InputCollection.Count;
                        item_list.total_number_of_items = 0;
                        let item_list.item_types: Object;
                        let i: number;
                        // With...
                        for (i = 1; (i <= item_list.num_item_types.item_types); i++) {
                            i.item_types(i).Width = 1;
                            i.id = 1;
                            if ((SteelClass == false)) {
                                item_list.item_types;
                                InputCollection(i).tLength.item_types(i).number_requested = InputCollection(i).Quantity;
                                InputCollection(i).tLength.item_types(i).Area = InputCollection(i).Quantity;
                                i.Height = InputCollection(i).Quantity;
                            }
                            else {
                                item_list.item_types;
                                InputCollection(i).Placement.item_types(i).Size = InputCollection(i).Size;
                                InputCollection(i).Qty.item_types(i).Placement = InputCollection(i).Size;
                                InputCollection(i).Length.item_types(i).number_requested = InputCollection(i).Size;
                                InputCollection(i).Length.item_types(i).Area = InputCollection(i).Size;
                                i.Height = InputCollection(i).Size;
                            }

                            item_list.item_types;
                            1.item_types(i).profit = item_list.item_types;
                            false.item_types(i).mandatory = item_list.item_types;
                            i.rotatable = item_list.item_types;
                            i.Height;
                            if ((solver_options.item_sort_criterion == 1)) {
                                item_list.item_types;
                                i.sort_criterion = item_list.item_types;
                                i.Area;
                            }
                            else if ((solver_options.item_sort_criterion == 2)) {
                                item_list.item_types;
                                i.sort_criterion = item_list.item_types;
                                (i.Width + item_list.item_types);
                                i.Height;
                            }
                            else if ((solver_options.item_sort_criterion == 3)) {
                                item_list.item_types;
                                i.sort_criterion = item_list.item_types;
                                i.Height;
                            }
                            else if ((solver_options.item_sort_criterion == 4)) {
                                item_list.item_types;
                                i.sort_criterion = item_list.item_types;
                                i.Width;
                            }

                            item_list.total_number_of_items = (item_list.total_number_of_items + item_list.item_types);
                            i.number_requested;
                        }

                        InitializeSolution((<solution_data>(solution)));
                        let i: number;
                        let j: number;
                        let k: number;
                        let l: number;
                        // With...
                        for (i = 1; (i <= bin_list.num_bin_types); i++) {
                            if ((bin_list.bin_types(i).mandatory >= 0)) {
                                (num_bins + bin_list.bin_types(i).number_available);
                                0.total_area_utilization = 0;
                                0.total_distance = 0;
                                0.total_area = 0;
                                false.net_profit = 0;
                                solution.feasible = 0;
                            }

                        }

                        let .: Object;
                        item_type_order(1, To, item_list.num_item_types);
                        for (i = 1; (i <= i); i++) {
                        }

                        let .: Object;
                        rotation_order(1, To, item_list.num_item_types, 1, To, 2);
                        for (i = 1; (i <= 1); i++) {
                        }

                        item_list.num_item_types.rotation_order(i, 1) = 1;
                        let .: Object;
                        first_cut_direction(1, To, item_list.num_item_types);
                        for (i = 1; (i <= 0); i++) {
                        }

                        let .: Object;
                        bin(1, To, ., num_bins);
                        for (i = 1; (i <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                        ); i++) {
                            num_bins;
                            let .: Object;
                            bin(i).items(1, To, (2 * item_list.total_number_of_items));
                            let .: Object;
                            bin(i).addition_points(1, To, (item_list.total_number_of_items + 1));
                            let .: Object;
                            bin(i).repack_item_count(1, To, item_list.total_number_of_items);
                        }

                        let .: Object;
                        unpacked_item_count(1, To, item_list.num_item_types);
                        l = 1;
                        for (i = 1; (i <= bin_list.num_bin_types); i++) {
                            if ((bin_list.bin_types(i).mandatory >= 0)) {
                                for (j = 1; (j <= 0); j++) {
                                    for (k = 1; (k <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                    ); k++) {
                                        bin(l).Height;
                                        i.bin(l).area_packed = 0;
                                        bin_list.bin_types(i).mandatory.bin(l).type_id = 0;
                                        bin_list.bin_types(i).cost.bin(l).mandatory = 0;
                                        bin_list.bin_types(i).Area.bin(l).cost = 0;
                                        bin_list.bin_types(i).Height.bin(l).Area = 0;
                                        bin_list.bin_types(i).Width.bin(l).Height = 0;
                                        bin_list.bin_types(i).number_available.bin(l).Width = 0;
                                    }

                                    for (k = 1; (k <= 0); k++) {
                                    }

                                    l = (l + 1);
                                }

                            }

                        }

                        for (i = 1; (i <= item_list.item_types[i].number_requested); i++) {
                        }

                        WriteSolution((<solution_data>(solution)), (<Collection>(SolvedCollection)), (<string>(InputType)));
                        let TrimPiece: clsTrim;
                        let CurrentBin: number;
                        let Member: clsMember;
                        let Span: clsMember;
                        Application.ScreenUpdating = false;
                        Application.Calculation = xlCalculationManual;
                        Application.EnableEvents = false;
                        let i: number;
                        let j: number;
                        let k: number;
                        let m: number;
                        let BinIndex: number;
                        let bin_index: number;
                        let swap_bin: bin_data;
                        // my vars
                        let BinName: string;
                        let ItemName: string;
                        let BinNumber: number;
                        // reset bin index
                        BinIndex = 1;
                        // sort the bins
                        // first ensure that simlar patterns occur together
                        for (i = 1; (i <= solution.num_bins); i++) {
                            solution.bin(i).area_packed = (solution.bin(i).area_packed
                                        * (solution.bin(i).Area * solution.bin(i).Area));
                            for (j = 1; (j <= solution.bin(i).item_cnt); j++) {
                                solution.bin(i).area_packed = (solution.bin(i).area_packed
                                            + (item_list.item_types[solution.bin(i).items(j).item_type].Area * item_list.item_types[solution.bin(i).items(j).item_type].Area));
                                Debug.Print;
                                item_list.item_types[solution.bin(i).items(j).item_type].Area;
                                Debug.Print;
                                item_list.item_types[solution.bin(i).items(j).item_type].id;
                                Debug.Print;
                                item_list.item_types[solution.bin(i).items(j).item_type].number_requested;
                            }

                        }

                        for (i = 1; (i <= solution.num_bins); i++) {
                            for (j = solution.num_bins; (j <= 2); j = (j + -1)) {
                                if (((solution.bin(j).type_id < solution.bin((j - 1)).type_id)
                                            || ((solution.bin(j).type_id == solution.bin((j - 1)).type_id)
                                            && (solution.bin(j).area_packed > solution.bin((j - 1)).area_packed)))) {
                                    swap_bin = solution.bin(j);
                                    solution.bin(j) = solution.bin((j - 1));
                                    solution.bin((j - 1)) = swap_bin;
                                }

                            }

                        }

                        if (((solution.feasible == false)
                                    && (InputType != "Roof Purlin"))) {
                            MsgBox;
                            ("Warning: Last solution returned by the solver does not satisfy all constraints. The parts were not optimized, partially optimized, or no optimization was possible. Please check "
                                        + (InputType + " parts."));
                        }

                        bin_index = 1;
                        // Expanded solution output - test
                        // With...
                        bin_index = 1;
                        CurrentBin = 0;
                        for (i = 1; (i <= solution.num_bins); i++) {
                            if (solution.bin) {
                                (i.item_cnt > 0);
                                Debug.Print;
                                ("Bin #: "
                                            + (i + (" " + solution.bin)));
                                i.Height;
                                Debug.Print;
                                ("item count: " + solution.bin);
                                i.item_cnt;
                                for (j = 1; (j <= solution.bin); j++) {
                                    i.item_cnt;
                                    if ((j > 1)) {
                                        (Debug.Print.bin(i).items(j).ne_y - solution.bin);
                                        i.items(j).sw_y;
                                    }
                                    else {
                                        Debug.Print.bin(i).items(j).ne_y;
                                    }

                                }

                            }

                        }

                        // With...
                        // condensed solution output
                        bin_index = 1;
                        CurrentBin = 0;
                        for (i = 1; (i <= bin_list.num_bin_types); i++) {
                            for (j = 1; (j <= bin_list.bin_types(i).number_available); j++) {
                                // name bins
                                switch (solution.bin) {
                                    case bin_index.Height:
                                        // '' steel
                                        break;
                                }

                                (25 * 12.bin(bin_index).bin_name) = "25'";
                                (30 * 12.bin(bin_index).bin_name) = "30'";
                                42.bin(bin_index).bin_name = "3'6""";
                                75.bin(bin_index).bin_name = "6'3""";
                                86.bin(bin_index).bin_name = "7'2""";
                                87.bin(bin_index).bin_name = "7'3""";
                                99.bin(bin_index).bin_name = "8'3""";
                                122.bin(bin_index).bin_name = "10'2""";
                                123.bin(bin_index).bin_name = "10'3""";
                                146.bin(bin_index).bin_name = "12'2""";
                                147.bin(bin_index).bin_name = "12'3""";
                                170.bin(bin_index).bin_name = "14'2""";
                                171.bin(bin_index).bin_name = "14'3""";
                                194.bin(bin_index).bin_name = "16'2""";
                                195.bin(bin_index).bin_name = "16'3""";
                                218.bin(bin_index).bin_name = "18'2""";
                                219.bin(bin_index).bin_name = "18'3""";
                                244.bin(bin_index).bin_name = "20'4""";
                                (bin_list.bin_types(i).mandatory >= 0);
                                // if items in the bin, add to the output bin number
                                if (solution.bin) {
                                    (bin_index.item_cnt != 0);
                                    BinNumber = (BinNumber + 1);
                                    // add previous trim piece
                                    // If BinNumber <> 1 Then SolvedCollection.Add TrimPiece
                                }

                                // check for items
                                for (k = 1; (k <= solution.bin); k++) {
                                    bin_index.item_cnt;
                                    if ((CurrentBin != BinNumber)) {
                                        if ((SteelMode == false)) {
                                            // new trim class
                                            TrimPiece = new clsTrim();
                                            TrimPiece.tMeasurement = solution.bin;
                                            bin_index.bin_name;
                                            TrimPiece.tLength = solution.bin;
                                            bin_index.Height;
                                            // girt whatever.placement = .bin(bin_index).placement
                                            switch (InputType) {
                                                case "Jamb":
                                                    TrimPiece.tType = "Jamb Trim";
                                                    break;
                                                case "Head":
                                                    TrimPiece.tType = "Head Trim W/ Kickout";
                                                    break;
                                            }

                                            TrimPiece.Quantity = 1;
                                            SolvedCollection.Add;
                                            TrimPiece;
                                        }
                                        else {
                                            // new member class
                                            Member = new clsMember();
                                            Member.Measurement = solution.bin;
                                            bin_index.bin_name;
                                            Member.Length = solution.bin;
                                            bin_index.Height;
                                            //  This is the total length (ie - 20' or 25' or 30')
                                            Member.Size = solution.bin;
                                            bin_index.items(1).Size;
                                            if ((InputType == "Roof Purlin")) {
                                                Member.Placement = (InputType + (" "
                                                            + (Member.Size + (" Span # " + BinIndex))));
                                            }
                                            else {
                                                Member.Placement = (Member.Size + (" Span # " + BinIndex));
                                            }

                                            BinIndex = (BinIndex + 1);
                                            // girt whatever.placement = .bin(bin_index).placement
                                            switch (InputType) {
                                                case "Girt":
                                                    Member.mType = "C Purlin";
                                                    break;
                                                case "Roof Purlin":
                                                    Member.mType = "C Purlin";
                                                    break;
                                                case "TS":
                                                    Member.mType = "Tube Steel";
                                                    break;
                                                case "IBeam":
                                                    Member.mType = "I-Beam";
                                                    break;
                                            }

                                            Member.Qty = 1;
                                            // debug output
                                            if ((SteelMode == true)) {
                                                for (m = 1; (m <= solution.bin); m++) {
                                                    bin_index.item_cnt;
                                                    if ((m > 1)) {
                                                        Span = new clsMember();
                                                        Span.Length = solution.bin;
                                                        (bin_index.items(m).ne_y - solution.bin);
                                                        bin_index.items(m).sw_y;
                                                        Span.Qty = 1;
                                                        Span.Placement = solution.bin;
                                                        bin_index.items(m).Placement;
                                                        Span.Size = solution.bin;
                                                        bin_index.items(m).Size;
                                                        Member.ComponentMembers.Add;
                                                        Span;
                                                        // Debug.Print "Added Member No.: " & k & ", Member length: "; .bin(bin_index).items(k).ne_y - .bin(bin_index).items(k - 1).ne_y & ", Placement: " & .bin(bin_index).items(k).Placement
                                                    }
                                                    else {
                                                        Span = new clsMember();
                                                        Span.Length = solution.bin;
                                                        bin_index.items(m).ne_y;
                                                        Span.Qty = 1;
                                                        Span.Placement = solution.bin;
                                                        bin_index.items(m).Placement;
                                                        Span.Size = solution.bin;
                                                        bin_index.items(m).Size;
                                                        Member.ComponentMembers.Add;
                                                        Span;
                                                        // Debug.Print "Bin # " & bin_index & ", Order Length: " & .bin(bin_index).bin_name
                                                        // Debug.Print "Added Member No.: " & k & ", Member length: "; .bin(bin_index).items(k).ne_y & ", Placement: " & .bin(bin_index).items(k).Placement
                                                    }

                                                }

                                            }

                                            SolvedCollection.Add;
                                            Member;
                                        }

                                        // update current bin
                                        CurrentBin = BinNumber;
                                    }

                                }

                                bin_index = (bin_index + 1);
                                j;
                            }

                            i;
                            // With...
                            Application.Calculation = xlCalculationAutomatic;
                            Application.EnableEvents = true;
                            ReadSolution((<solution_data>(solution)));
                            Application.ScreenUpdating = false;
                            Application.Calculation = xlCalculationManual;
                            let i: number;
                            let j: number;
                            let k: number;
                            let l: number;
                            let bin_index: number;
                            let item_type_index: number;
                            let offset: number;
                            offset = 0;
                            bin_index = 1;
                            // With...
                            for (i = 1; (i <= bin_list.num_bin_types); i++) {
                                for (j = 1; (j <= bin_list.bin_types(i).number_available); j++) {
                                    if ((bin_list.bin_types(i).mandatory >= 0)) {
                                        // With...
                                        bin_index;
                                        l = Cells(4, (offset + 17)).Value;
                                        for (k = 1; (k <= l); k++) {
                                            if ((IsNumeric(Cells((5 + k), (offset + 14)).Value) == true)) {
                                                solution.bin.item_cnt = (solution.bin.item_cnt + 1);
                                                item_type_index = Cells((5 + k), (offset + 14)).Value;
                                                Cells((5 + k), (offset + 3)).Value.items(solution.bin.item_cnt).sw_y = Cells((5 + k), (offset + 4)).Value;
                                                item_type_index.items(solution.bin.item_cnt).sw_x = Cells((5 + k), (offset + 4)).Value;
                                                (solution.unpacked_item_count(item_type_index) - 1.items(solution.bin.item_cnt).item_type) = Cells((5 + k), (offset + 4)).Value;
                                                solution.unpacked_item_count(item_type_index) = Cells((5 + k), (offset + 4)).Value;
                                                if ((ThisWorkbook.Worksheets("3.Solution").Cells[(5 + k), (offset + 5)].Value == "Yes")) {
                                                    solution.bin.items;
                                                    true.items(solution.bin.item_cnt).ne_x = solution.bin.items;
                                                    solution.bin.item_cnt.rotated = solution.bin.items;
                                                    (solution.bin.item_cnt.sw_x + item_list.item_types[item_type_index].Height.items(solution.bin.item_cnt).ne_y) = solution.bin.items;
                                                    (solution.bin.item_cnt.sw_y + item_list.item_types[item_type_index].Width);
                                                }
                                                else {
                                                    solution.bin.items;
                                                    false.items(solution.bin.item_cnt).ne_x = solution.bin.items;
                                                    solution.bin.item_cnt.rotated = solution.bin.items;
                                                    (solution.bin.item_cnt.sw_x + item_list.item_types[item_type_index].Width.items(solution.bin.item_cnt).ne_y) = solution.bin.items;
                                                    (solution.bin.item_cnt.sw_y + item_list.item_types[item_type_index].Height);
                                                }

                                                if ((instance.guillotine_cuts == true)) {
                                                    if ((ThisWorkbook.Worksheets("3.Solution").Cells[(5 + k), (offset + 6)].Value == solution.bin.items)) {
                                                        solution.bin.item_cnt.sw_x;
                                                        solution.bin.items;
                                                        0.items(solution.bin.item_cnt).max_x = Cells((5 + k), (offset + 8)).Value;
                                                        solution.bin.item_cnt.first_cut_direction = Cells((5 + k), (offset + 8)).Value;
                                                    }
                                                    else {
                                                        solution.bin.items;
                                                        1.items(solution.bin.item_cnt).max_y = Cells((5 + k), (offset + 9)).Value;
                                                        solution.bin.item_cnt.first_cut_direction = Cells((5 + k), (offset + 9)).Value;
                                                    }

                                                }

                                                solution.bin.area_packed = (solution.bin.area_packed + item_list.item_types[item_type_index].Area);
                                                if ((solution.bin.item_cnt == 1)) {
                                                    solution.net_profit = (solution.net_profit
                                                                + (item_list.item_types[item_type_index].profit - solution.bin.cost));
                                                }
                                                else {
                                                    solution.net_profit = (solution.net_profit + item_list.item_types[item_type_index].profit);
                                                }

                                            }

                                        }

                                        bin_index = (bin_index + 1);
                                    }

                                    offset = (offset + offset_constant);
                                }

                            }

                            Application.ScreenUpdating = true;
                            Application.Calculation = xlCalculationAutomatic;
                            BPP_Solver((<Collection>(SolvedCollection)), (<Collection>(InputCollection)), (<string>(InputType)), Optional, (<string>(FOType)), Optional, (<string>(Wall)));
                            let i: number;
                            let AvailableBins: number;
                            let result: number;
                            let ItemCount: number;
                            let TrimPiece: clsTrim;
                            let TotalTrimLength: number;
                            let TotalMemberLength: number;
                            let Member: clsMember;
                            // determine total item count
                            if (((InputType != "Girt")
                                        && ((InputType != "Roof Purlin")
                                        && ((InputType != "TS")
                                        && (InputType != "IBeam"))))) {
                                for (TrimPiece in InputCollection) {
                                    ItemCount = (ItemCount + TrimPiece.Quantity);
                                    TotalTrimLength = (TotalTrimLength
                                                + (TrimPiece.tLength * TrimPiece.Quantity));
                                }

                                Debug.Print;
                                (FOType + (" "
                                            + (InputType + (" " + TotalTrimLength))));
                                SteelMode = false;
                            }
                            else {
                                SteelMode = true;
                                for (Member in InputCollection) {
                                    ItemCount = (ItemCount + Member.Qty);
                                    TotalMemberLength = (TotalMemberLength
                                                + (Member.Length * Member.Qty));
                                }

                                Debug.Print;
                                (Wall + (" "
                                            + (InputType + (" " + TotalMemberLength))));
                            }

                            // Instance data variables
                            Application.ScreenUpdating = false;
                            Application.Calculation = xlCalculationManual;
                            let WorksheetExists: boolean;
                            let reply: number;
                            Application.EnableCancelKey = xlErrorHandler;
                            // TODO: On Error GoTo Warning!!!: The statement is not translatable
                            // ''''''''''''''''''''''''''//// Set Initial Solver Options ///
                            // With...
                            // '' Sort Criteria: 1 = Area, 2 = Circumference, 3 = height, 4 = Width
                            // '' Show_Progress property = Display progress in status bar
                            // '' CPU_Time Limit: Run Time in Seconds. Minimum: 1 second/item
                            3.show_progress = true;
                            solver_options.item_sort_criterion = true;
                            if (((FOType == "OHDoors")
                                        || ((FOType == "MiscFOs")
                                        || (FOType == "")))) {
                                solver_options.CPU_time_limit = (ItemCount * 5);
                            }
                            else if ((InputType == "Girt")) {
                                solver_options.CPU_time_limit = (ItemCount / 5);
                            }
                            else {
                                solver_options.CPU_time_limit = (ItemCount / 5);
                            }

                            // ''''''''''''''''''''''''''//// Set up "Items" (The Trim Lengths to be Ordered) ///
                            if (((InputType == "Girt")
                                        || ((InputType == "Roof Purlin")
                                        || ((InputType == "TS")
                                        || (InputType == "IBeam"))))) {
                                GetItemData(InputCollection, true);
                            }
                            else {
                                GetItemData(InputCollection);
                            }

                            // ''''''''''''''''''''''''''//// Set up "Bins" (Stock Standard Trim Sizes" ///
                            //  Calculate Max Possible Bins Needed
                            AvailableBins = 5;
                            // Placeholder
                            // With...
                            // number of stock trim types
                            if ((InputType == "Jamb")) {
                                bin_list.num_bin_types = 7;
                            }
                            else if ((InputType == "Head")) {
                                bin_list.num_bin_types = 11;
                            }
                            else if ((InputType == "Girt")) {
                                bin_list.num_bin_types = 3;
                            }
                            else if ((InputType == "Roof Purlin")) {
                                bin_list.num_bin_types = 3;
                            }
                            else if ((InputType == "TS")) {
                                bin_list.num_bin_types = 3;
                            }
                            else if ((InputType == "IBeam")) {
                                bin_list.num_bin_types = 9;
                            }

                            let .: Object;
                            bin_types(1, To, bin_list.num_bin_types);
                            for (i = 1; (i <= bin_list.num_bin_types.bin_types); i++) {
                                i.bin_types(i).Width = 1;
                                i.type_id = 1;
                                // Mandatory Values: 1 = Must use; 0 = May use; -1 = Don't use
                                bin_list.bin_types;
                                i.mandatory = 0;
                                // Number available: needs to be calculated to optimize (I think)
                                if ((InputType == "Jamb")) {
                                    switch (i) {
                                    }

                                    1.bin_types(i).Height = "7'2""";
                                    122.bin_types(i).bin_name = "10'2""";
                                    2.bin_types(i).Height = "10'2""";
                                    146.bin_types(i).bin_name = "12'2""";
                                    3.bin_types(i).Height = "12'2""";
                                    170.bin_types(i).bin_name = "14'2""";
                                    4.bin_types(i).Height = "14'2""";
                                    194.bin_types(i).bin_name = "16'2""";
                                    5.bin_types(i).Height = "16'2""";
                                    218.bin_types(i).bin_name = "18'2""";
                                    6.bin_types(i).Height = "18'2""";
                                    244.bin_types(i).bin_name = "20'4""";
                                    7.bin_types(i).Height = "20'4""";
                                }

                                InputType = "Head";
                                switch (i) {
                                }

                                1.bin_types(i).Height = "3'6""";
                                75.bin_types(i).bin_name = "6'3""";
                                2.bin_types(i).Height = "6'3""";
                                87.bin_types(i).bin_name = "7'3""";
                                3.bin_types(i).Height = "7'3""";
                                99.bin_types(i).bin_name = "8'3""";
                                4.bin_types(i).Height = "8'3""";
                                122.bin_types(i).bin_name = "10'2""";
                                5.bin_types(i).Height = "10'2""";
                                123.bin_types(i).bin_name = "10'3""";
                                6.bin_types(i).Height = "10'3""";
                                147.bin_types(i).bin_name = "12'3""";
                                7.bin_types(i).Height = "12'3""";
                                171.bin_types(i).bin_name = "14'3""";
                                8.bin_types(i).Height = "14'3""";
                                195.bin_types(i).bin_name = "16'3""";
                                9.bin_types(i).Height = "16'3""";
                                219.bin_types(i).bin_name = "18'3""";
                                10.bin_types(i).Height = "18'3""";
                                244.bin_types(i).bin_name = "20'4""";
                                11.bin_types(i).Height = "20'4""";
                                ("Girt" | InputType) = "Roof Purlin";
                                InputType = "Roof Purlin";
                                switch (i) {
                                }

                                1.bin_types(i).Height = "20'";
                                (25 * 12.bin_types(i).bin_name) = "25'";
                                2.bin_types(i).Height = "25'";
                                (30 * 12.bin_types(i).bin_name) = "30'";
                                3.bin_types(i).Height = "30'";
                                InputType = "TS";
                                switch (i) {
                                }

                                1.bin_types(i).Height = "20'";
                                (24 * 12.bin_types(i).bin_name) = "24'";
                                2.bin_types(i).Height = "24'";
                                (40 * 12.bin_types(i).bin_name) = "40'";
                                3.bin_types(i).Height = "40'";
                                InputType = "IBeam";
                                switch (i) {
                                }

                                1.bin_types(i).Height = "20'";
                                (25 * 12.bin_types(i).bin_name) = "25'";
                                2.bin_types(i).Height = "25'";
                                (30 * 12.bin_types(i).bin_name) = "30'";
                                3.bin_types(i).Height = "30'";
                                (35 * 12.bin_types(i).bin_name) = "35'";
                                4.bin_types(i).Height = "35'";
                                (40 * 12.bin_types(i).bin_name) = "40'";
                                5.bin_types(i).Height = "40'";
                                (45 * 12.bin_types(i).bin_name) = "45'";
                                6.bin_types(i).Height = "45'";
                                (50 * 12.bin_types(i).bin_name) = "50'";
                                7.bin_types(i).Height = "50'";
                                (55 * 12.bin_types(i).bin_name) = "55'";
                                8.bin_types(i).Height = "55'";
                                (60 * 12.bin_types(i).bin_name) = "60'";
                                9.bin_types(i).Height = "60'";
                                if (bin_list.bin_types) {
                                    i.Area = bin_list.bin_types;
                                    i.Height.bin_types(i).cost = bin_list.bin_types;
                                    i.Height;
                                    if (((InputType != "Girt")
                                                && ((InputType != "Roof Purlin")
                                                && ((InputType != "TS")
                                                && (InputType != "IBeam"))))) {
                                        bin_list.bin_types;
                                        i.number_available = Application.WorksheetFunction.RoundUp((TotalTrimLength / bin_list.bin_types), i.Height, 0);
                                    }
                                    else {
                                        bin_list.bin_types;
                                        i.number_available = Application.WorksheetFunction.RoundUp((TotalMemberLength / bin_list.bin_types), i.Height, 0);
                                    }

                                    // lower number available if over item count
                                    if (bin_list.bin_types) {
                                        (i.number_available > ItemCount);
                                        bin_list.bin_types;
                                        i.number_available = ItemCount;
                                        Debug.Print;
                                        (FOType
                                                    + (Wall + (", "
                                                    + (InputType + (" - " + bin_list.bin_types)))));
                                        (i.Height + (" Available:" + bin_list.bin_types));
                                        i.number_available;
                                        i;
                                    }

                                    // ''''''''''''''''''''''''''//// Set "Instance" Options; Determine the Max Potential Profit - The "Upper Bound" ///
                                    result = 0;
                                    //  for mandatory bins and possible items, determine the total possible profit for items and the total cost for bins
                                    // With...
                                    for (i = 1; (i <= item_list.num_item_types); i++) {
                                        if (item_list.item_types) {
                                            (i.mandatory != -1);
                                            result = (result
                                                        + (item_list.item_types[i].profit * item_list.item_types));
                                            i.number_requested;
                                        }

                                    }

                                    // set max possible profit
                                    instance.global_upper_bound = result;
                                    // ''' Setting Compatability Sheets to False- everything is compatible (for now)
                                    //  note: This could change if I want to run the program for head and jamb trim at the same time
                                    instance.item_item_compatibility_worksheet = false;
                                    instance.bin_item_compatibility_worksheet = false;
                                    instance.guillotine_cuts = false;
                                    SortItems();
                                    // ''''''''''''''''''''''''''//// Start Solution Setup Stuff ///
                                    let incumbent: solution_data;
                                    InitializeSolution(incumbent);
                                    let best_known: solution_data;
                                    InitializeSolution(best_known);
                                    best_known = incumbent;
                                    let iteration: number;
                                    let j: number;
                                    let k: number;
                                    let l: number;
                                    let start_time: Date;
                                    let end_time: Date;
                                    let continue_flag: boolean;
                                    // infeasibility check
                                    let infeasibility_count: number;
                                    let infeasibility_string: string;
                                    // checks to see that all items can fit, etc
                                    // Call FeasibilityCheckData(infeasibility_count, infeasibility_string)
                                    start_time = Timer;
                                    end_time = Timer;
                                    // constructive phase
                                    if ((solver_options.show_progress == true)) {
                                        Application.ScreenUpdating = true;
                                        Application.StatusBar = "Constructive phase...";
                                        Application.ScreenUpdating = false;
                                    }
                                    else {
                                        Application.ScreenUpdating = true;
                                        Application.StatusBar = "LNS algorithm running...";
                                        Application.ScreenUpdating = false;
                                    }

                                    SortBins(incumbent);
                                    for (i = 1; (i <= incumbent.num_bins); i++) {
                                        for (j = 1; (j <= item_list.num_item_types); j++) {
                                            continue_flag = true;
                                            while (((incumbent.unpacked_item_count(incumbent.item_type_order(j)) > 0)
                                                        && (continue_flag == true))) {
                                                continue_flag = AddItemToBin(incumbent, i, incumbent.item_type_order(j), 1);
                                            }

                                        }

                                        incumbent.feasible = true;
                                        for (j = 1; (j <= item_list.num_item_types); j++) {
                                            if (((incumbent.unpacked_item_count(j) > 0)
                                                        && (item_list.item_types[j].mandatory == 1))) {
                                                incumbent.feasible = false;
                                                break;
                                            }

                                        }

                                        CalculateDistance(incumbent);
                                        if ((((incumbent.feasible == true)
                                                    && (best_known.feasible == false))
                                                    || (((incumbent.feasible == false)
                                                    && ((best_known.feasible == false)
                                                    && (incumbent.total_area
                                                    > (best_known.total_area + epsilon))))
                                                    || (((incumbent.feasible == true)
                                                    && ((best_known.feasible == true)
                                                    && (incumbent.net_profit
                                                    > (best_known.net_profit + epsilon))))
                                                    || ((incumbent.feasible == true)
                                                    && ((best_known.feasible == true)
                                                    && ((incumbent.net_profit
                                                    > (best_known.net_profit - epsilon))
                                                    && (incumbent.total_area
                                                    < (best_known.total_area - epsilon))))))))) {
                                            best_known = incumbent;
                                        }

                                    }

                                    end_time = Timer;
                                    // MsgBox "Constructive phase result: " & best_known.net_profit & " time: " & end_time - start_time
                                    // improvement phase
                                    iteration = 0;
                                    for (
                                    ; ((end_time - start_time)
                                                < solver_options.CPU_time_limit);
                                    ) {
                                        DoEvents;
                                        if (((solver_options.show_progress == true)
                                                    && ((iteration % 100)
                                                    == 0))) {
                                            Application.ScreenUpdating = true;
                                            if (SteelMode) {
                                                if ((best_known.feasible == true)) {
                                                    Application.StatusBar = ("Starting "
                                                                + (FOType + (" "
                                                                + (InputType + (" " + ("Iteration #" + iteration))))));
                                                    // & ". Best net profit found so far: " & best_known.net_profit ' & " TAU: " & best_known.total_area_utilization
                                                }
                                                else {
                                                    Application.StatusBar = ("Starting iteration "
                                                                + (iteration + ". Best net profit found so far: N/A"));
                                                }

                                            }
                                            else if ((best_known.feasible == true)) {
                                                Application.StatusBar = ("Starting "
                                                            + (FOType + (" "
                                                            + (InputType + (" Trim " + ("Iteration #" + iteration))))));
                                                // & ". Best net profit found so far: " & best_known.net_profit ' & " TAU: " & best_known.total_area_utilization
                                            }
                                            else {
                                                Application.StatusBar = ("Starting iteration "
                                                            + (iteration + ". Best net profit found so far: N/A"));
                                            }

                                            Application.ScreenUpdating = false;
                                        }

                                        if ((Rnd() < 0.5)) {
                                            // Rnd() < ((end_time - start_time) / solver_options.CPU_time_limit) ^ 2 Then
                                            incumbent = best_known;
                                        }

                                        PerturbSolution(incumbent);
                                        SortBins(incumbent);
                                        // With...
                                        for (i = 1; (i <= incumbent.num_bins); i++) {
                                            for (j = 1; (j <= item_list.num_item_types); j++) {
                                                continue_flag = true;
                                                while ((incumbent.unpacked_item_count[incumbent.item_type_order, j] > 0)) {
                                                    (continue_flag == true);
                                                    continue_flag = AddItemToBin(incumbent, i, incumbent.item_type_order, j, 1);
                                                }

                                            }

                                            j.feasible = true;
                                            for (j = 1; (j <= item_list.num_item_types); j++) {
                                                if ((incumbent.unpacked_item_count[j] > 0)) {
                                                    (item_list.item_types[j].mandatory == 1);
                                                    incumbent.feasible = false;
                                                    break;
                                                }

                                            }

                                            CalculateDistance(incumbent);
                                            if ((((incumbent.feasible == true)
                                                        && (best_known.feasible == false))
                                                        || (((incumbent.feasible == false)
                                                        && ((best_known.feasible == false)
                                                        && (incumbent.total_area
                                                        > (best_known.total_area + epsilon))))
                                                        || ((((incumbent.feasible == false)
                                                        && ((best_known.feasible == false)
                                                        && (incumbent.total_area
                                                        > (best_known.total_area - epsilon))))
                                                        && (incumbent.total_distance
                                                        < (best_known.total_distance - epsilon)))
                                                        || (((incumbent.feasible == true)
                                                        && ((best_known.feasible == true)
                                                        && (incumbent.net_profit
                                                        > (best_known.net_profit + epsilon))))
                                                        || (((incumbent.feasible == true)
                                                        && ((best_known.feasible == true)
                                                        && ((incumbent.net_profit
                                                        > (best_known.net_profit - epsilon))
                                                        && (incumbent.total_area
                                                        < (best_known.total_area - epsilon)))))
                                                        || ((((incumbent.feasible == true)
                                                        && ((best_known.feasible == true)
                                                        && ((incumbent.net_profit
                                                        > (best_known.net_profit - epsilon))
                                                        && (incumbent.total_area
                                                        < (best_known.total_area + epsilon)))))
                                                        && (incumbent.total_area_utilization > best_known.total_area_utilization))
                                                        || (((incumbent.feasible == true)
                                                        && ((best_known.feasible == true)
                                                        && ((incumbent.net_profit
                                                        > (best_known.net_profit - epsilon))
                                                        && (incumbent.total_area
                                                        < (best_known.total_area + epsilon)))))
                                                        && ((incumbent.total_area_utilization >= best_known.total_area_utilization)
                                                        && (incumbent.total_distance
                                                        < (best_known.total_distance - epsilon))))))))))) {
                                                best_known = incumbent;
                                            }

                                        }

                                        iteration = (iteration + 1);
                                        end_time = Timer;
                                    }

                                    // MsgBox "Iterations performed: " & iteration
                                    /* Warning! Labeled Statements are not Implemented */if ((best_known.feasible == true)) {
                                        WriteSolution(best_known, SolvedCollection, InputType);
                                    }
                                    else if ((infeasibility_count > 0)) {
                                        WriteSolution(best_known, SolvedCollection, InputType);
                                    }
                                    else {
                                        // reply = MsgBox("The best found solution after " & iteration & " LNS iterations does not satisfy all constraints. Do you want to overwrite the current solution with the best found solution?", vbYesNo, "BPP Spreadsheet Solver")
                                        reply = System.Windows.Forms.MessageBoxButtons.Yes;
                                        if ((reply == System.Windows.Forms.MessageBoxButtons.Yes)) {
                                            WriteSolution(best_known, SolvedCollection, InputType);
                                        }

                                    }

                                    // Erase the data
                                    item_list.item_types;
                                    bin_list.bin_types;
                                    compatibility_list.item_to_item;
                                    compatibility_list.bin_to_item;
                                    for (i = 1; (i <= incumbent.num_bins); i++) {
                                        incumbent.bin(i).items;
                                    }

                                    incumbent.bin;
                                    incumbent.unpacked_item_count;
                                    for (i = 1; (i <= best_known.num_bins); i++) {
                                        best_known.bin(i).items;
                                    }

                                    best_known.bin;
                                    best_known.unpacked_item_count;
                                    Application.StatusBar = false;
                                    Application.ScreenUpdating = true;
                                    Application.Calculation = xlCalculationAutomatic;
                                }

                                SortItems();
                                let i: number;
                                let j: number;
                                let swap_item_type: item_type_data;
                                if ((item_list.num_item_types > 1)) {
                                    for (i = 1; (i <= item_list.num_item_types); i++) {
                                        for (j = item_list.num_item_types; (j <= 2); j = (j + -1)) {
                                            // checks: 1. if previous item is mandatory and current item isn't isnt
                                            // 2. If items are equally mandatory AND current item has a sort criteria field value higher than the previous item
                                            // 3. both not mandatory and current is more profitable than the previous
                                            // overall, sub sorts items in ascending importance/length
                                            if (((item_list.item_types[j].mandatory > item_list.item_types[(j - 1)].mandatory)
                                                        || (((item_list.item_types[j].mandatory == 1)
                                                        && ((item_list.item_types[(j - 1)].mandatory == 1)
                                                        && (item_list.item_types[j].sort_criterion > item_list.item_types[(j - 1)].sort_criterion)))
                                                        || ((item_list.item_types[j].mandatory == 0)
                                                        && ((item_list.item_types[(j - 1)].mandatory == 0)
                                                        && ((item_list.item_types[j].profit / item_list.item_types[j].Area)
                                                        > (item_list.item_types[(j - 1)].profit / item_list.item_types[(j - 1)].Area))))))) {
                                                swap_item_type = item_list.item_types[j];
                                                item_list.item_types[j] = item_list.item_types[(j - 1)];
                                                item_list.item_types[(j - 1)] = swap_item_type;
                                            }

                                        }

                                    }

                                }

                                //     For i = 1 To item_list.num_item_types
                                //        MsgBox item_list.item_types(i).id & " " & item_list.item_types(i).width & " " & item_list.item_types(i).height & " "
                                //     Next i
                                CalculateDistance((<solution_data>(solution)));
                                let i: number;
                                let j: number;
                                let k: number;
                                let l: number;
                                let item_flag: boolean;
                                let bin_count: number;
                                let penalty: number;
                                // for not fitting an item type into a single bin
                                if ((instance.guillotine_cuts == true)) {
                                    // With...
                                    for (j = 1; (j <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                    ); j++) {
                                        num_bins;
                                        // With...
                                        bin(j);
                                        for (k = 1; (k <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                        ); k++) {
                                            item_cnt;
                                            items(k).cut_length;
                                        }

                                    }

                                }
                                else {
                                    penalty = 1000;
                                    //  perhaps find a better value here?
                                    // With...
                                    for (j = 1; (j <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                    ); j++) {
                                        num_bins;
                                        // With...
                                        bin(j);
                                        for (k = 1; (k <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                        ); k++) {
                                            item_cnt;
                                            for (l = (k + 1); (l <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                            ); l++) {
                                                item_cnt;
                                                items(l).item_type;
                                                solution.total_distance = (solution.total_distance
                                                            + (Abs(., items(kUnknown.ne_x+, ., items(kUnknown.sw_x-, ., items(lUnknown.ne_x-, ., items(l).sw_x) + Abs(., items(kUnknown.ne_y+, ., items(kUnknown.sw_y-, ., items(lUnknown.ne_y-, ., items(l).sw_y)));
                                                l;
                                            }

                                            k;
                                            // With...
                                            j;
                                            for (i = 1; (i <= item_list.num_item_types); i++) {
                                                bin_count = 0;
                                                for (j = 1; (j <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                                ); j++) {
                                                    num_bins;
                                                    // With...
                                                    bin(j);
                                                    item_flag = false;
                                                    for (k = 1; (k <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                                    ); k++) {
                                                        item_cnt;
                                                        items(k).item_type = i;
                                                        item_flag = true;
                                                        break;
                                                        k;
                                                        if ((item_flag == true)) {
                                                            bin_count = (bin_count + 1);
                                                        }

                                                    }

                                                    j;
                                                    solution.total_distance = (solution.total_distance
                                                                + (penalty
                                                                * (bin_count * bin_count)));
                                                    i;
                                                    solution.total_area_utilization = 0;
                                                    for (j = 1; (j <= // TODO: Warning!!!! NULL EXPRESSION DETECTED...
                                                    ); j++) {
                                                        bin(j).Area;
                                                        2;
                                                    }

                                                    // With...
                                                    // ' ribbon calls and tab activation
                                                    //
                                                    // #If Win32 Or Win64 Or (MAC_OFFICE_VERSION >= 15) Then
                                                    //
                                                    // Sub BPP_Solver_ribbon_call(control As IRibbonControl)
                                                    //     Call BPP_Solver
                                                    // End Sub
                                                    //
                                                    //
                                                    // #End If
                                                }

                                            }

                                        }

                                    }

                                }

                            }

                        }

                    }

                }

            }

        }

    }
