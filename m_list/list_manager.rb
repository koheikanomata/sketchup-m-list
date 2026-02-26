# frozen_string_literal: true
# M List - リスト操作

module EstimateAuto
  def self.list_items
    @list_items ||= []
  end

  def self.clear_list
    @list_items = []
  end

  def self.persist_list_ids
    return unless model
    ids = list_items.map { |it| it[:id].to_s }
    model.set_attribute(DICT, LIST_IDS_KEY, ids.to_json)
  end

  def self.find_entity_by_pid(pid)
    return nil unless model
    result = model.find_entity_by_persistent_id(pid.to_i)
    result.is_a?(Array) ? result.first : result
  rescue StandardError
    nil
  end

  def self.find_item_in_tree(items, target_id)
    items.each do |it|
      return it if it[:id].to_s == target_id.to_s
      found = find_item_in_tree(it[:children] || [], target_id)
      return found if found
    end
    nil
  end

  def self.build_list_item(e)
    pid = e.respond_to?(:persistent_id) ? e.persistent_id.to_s : object_id.to_s
    detail = (e.respond_to?(:bounds) ? size_string(e) : nil) || "—"
    {
      id: pid,
      type: entity_type_label(e),
      tag: entity_tag(e),
      name: "",
      detail: detail,
      price: 0,
      memo: "",
      material_id: "",
      finish_id: "",
      material_name: ""
    }
  end

  def self.build_list_item_tree(e)
    item = build_list_item(e)
    children = collect_children(e).map { |child| build_list_item_tree(child) }
    item[:children] = children if children.any?
    if children.empty?
      item[:shape_type] = determine_shape_type(e)
    else
      item[:shape_type] = "nested"
    end
    item
  end

  def self.replace_item_in_tree(items, target_id, new_item)
    items.each_with_index do |it, i|
      if it[:id].to_s == target_id
        items[i] = new_item
        return true
      end
      next unless it[:children]&.any?
      return true if replace_item_in_tree(it[:children], target_id, new_item)
    end
    false
  end

  def self.update_item_in_tree(items, target_id, key, value)
    items.each do |it|
      if it[:id].to_s == target_id.to_s
        it[key] = value
        return true
      end
      return true if update_item_in_tree(it[:children] || [], target_id, key, value)
    end
    false
  end

  def self.propagate_nested_prices!(items)
    items.each do |it|
      propagate_nested_prices!(it[:children] || []) if it[:children]&.any?
      if it[:children]&.any?
        it[:price] = it[:children].sum { |c| (c[:price] || 0).to_f }.round
      end
    end
  end

  def self.merge_for_update(old_item, new_item, csv_data, entity = nil)
    id = new_item[:id].to_s
    csv_row = csv_data&.dig(id)
    if entity && entity.respond_to?(:get_attribute)
      new_item[:material_id] = entity.get_attribute(DICT, ATTR_MATERIAL_ID, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row&.dig(:material_id) || old_item&.dig(:material_id) || "")
      new_item[:finish_id] = entity.get_attribute(DICT, ATTR_FINISH_ID, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row&.dig(:finish_id) || old_item&.dig(:finish_id) || "")
      new_item[:material_name] = entity.get_attribute(DICT, ATTR_MATERIAL_NAME, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row&.dig(:material_name) || old_item&.dig(:material_name) || "")
      new_item[:name] = entity.get_attribute(DICT, ATTR_NAME, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row ? csv_row[:name].to_s.strip : (old_item&.dig(:name).to_s || ""))
      price_attr = entity.get_attribute(DICT, ATTR_PRICE, -999)
      new_item[:price] = ((price_attr != -999 && price_attr >= 0 ? price_attr.to_f : nil) || (csv_row ? csv_row[:price].to_f : (old_item&.dig(:price).to_f || 0))).to_f.round
      new_item[:memo] = entity.get_attribute(DICT, ATTR_MEMO, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row ? csv_row[:memo].to_s.strip : (old_item&.dig(:memo).to_s || ""))
      hidden_val = entity.get_attribute(DICT, ATTR_HIDDEN, false)
      new_item[:hidden] = hidden_val == true || hidden_val.to_s.downcase == "true" || hidden_val == 1
      unless new_item[:children]&.any?
        attr_shape = entity.get_attribute(DICT, ATTR_SHAPE_TYPE, "").to_s.strip
        new_item[:shape_type] = attr_shape if %w[beam board freeform other].include?(attr_shape)
      end
    else
      new_item[:material_id] = csv_row ? csv_row[:material_id].to_s.strip : (old_item&.dig(:material_id).to_s || "")
      new_item[:finish_id] = csv_row ? csv_row[:finish_id].to_s.strip : (old_item&.dig(:finish_id).to_s || "")
      new_item[:material_name] = csv_row ? csv_row[:material_name].to_s.strip : (old_item&.dig(:material_name).to_s || "")
      new_item[:name] = csv_row ? csv_row[:name].to_s.strip : (old_item&.dig(:name).to_s || "")
      new_item[:price] = (csv_row ? csv_row[:price].to_f : (old_item&.dig(:price).to_f || 0)).round
      new_item[:memo] = csv_row ? csv_row[:memo].to_s.strip : (old_item&.dig(:memo).to_s || "")
      new_item[:hidden] = old_item&.dig(:hidden) ? true : false
    end
    if %w[beam board].include?(new_item[:shape_type]) && new_item[:price].to_f <= 0
      calc = compute_price_for_item(new_item, entity)
      new_item[:price] = calc.round if calc && calc > 0
    end
    return unless new_item[:children]&.any?
    new_item[:children].each do |new_c|
      old_c = (old_item&.dig(:children) || []).find { |o| o[:id].to_s == new_c[:id].to_s }
      child_entity = entity && find_entity_by_pid(new_c[:id])
      merge_for_update(old_c, new_c, csv_data, child_entity)
    end
  end

  def self.apply_edits_from_payload(edits_json)
    return if edits_json.to_s.strip.empty?
    edits = JSON.parse(edits_json)
    return unless model
    model.start_operation("編集を反映", true)
    begin
      edits.each do |id_str, data|
        next unless id_str && !id_str.to_s.empty?
        h = data.is_a?(Hash) ? data : {}
        update_item_in_tree(list_items, id_str, :name, (h["name"] || h[:name] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :price, ((h["price"] || h[:price] || 0).to_f).round)
        update_item_in_tree(list_items, id_str, :memo, (h["memo"] || h[:memo] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :material_id, (h["material_id"] || h[:material_id] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :finish_id, (h["finish_id"] || h[:finish_id] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :material_name, (h["material_name"] || h[:material_name] || "").to_s.strip)
        entity = find_entity_by_pid(id_str)
        if entity && entity.respond_to?(:set_attribute)
          entity.set_attribute(DICT, ATTR_NAME, (h["name"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_PRICE, ((h["price"] || 0).to_f).round)
          entity.set_attribute(DICT, ATTR_MEMO, (h["memo"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_MATERIAL_ID, (h["material_id"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_FINISH_ID, (h["finish_id"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_MATERIAL_NAME, (h["material_name"] || "").to_s.strip)
        end
      end
      propagate_nested_prices!(list_items)
      model.commit_operation
    rescue StandardError
      model.abort_operation
    end
  rescue StandardError
  end
end
