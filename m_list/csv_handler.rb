# frozen_string_literal: true
# M List - CSV 読み書き・同期

module EstimateAuto
  def self.load_csv_data(path)
    return {} unless path && File.exist?(path)
    data = {}
    CSV.foreach(path, headers: true, encoding: "UTF-8") do |row|
      id = row["id"]&.to_s&.strip
      next if id.nil? || id.empty?
      data[id] = {
        name: (row["名前"] || row["name"] || "").to_s.strip,
        price: (row["値段"] || row["price"] || 0).to_f,
        memo: (row["メモ"] || row["memo"] || "").to_s.strip,
        material_id: (row["material_id"] || "").to_s.strip,
        finish_id: (row["finish_id"] || "").to_s.strip,
        material_name: (row["material_name"] || "").to_s.strip
      }
    end
    data
  rescue StandardError
    {}
  end

  def self.shape_type_label(st)
    case st
    when "nested" then "入れ子"
    when "beam" then "棒"
    when "board" then "面状"
    when "freeform" then "自由形状"
    when "other" then "その他"
    else "—"
    end
  end

  def self.flatten_items_for_csv(items, _indent_level)
    result = []
    items.each do |it|
      next if it[:hidden]
      shape = (it[:children]&.any? ? "入れ子" : shape_type_label(it[:shape_type]))
      result << [
        (it[:id] || "").to_s, (it[:tag] || "").to_s, (it[:name] || "").to_s,
        (it[:type] || "").to_s, shape, (it[:detail] || "").to_s,
        (it[:material_id] || "").to_s, (it[:finish_id] || "").to_s, (it[:material_name] || "").to_s,
        (it[:price] || 0).to_f.round, (it[:memo] || "").to_s
      ]
      result.concat(flatten_items_for_csv(it[:children] || [], 0))
    end
    result
  end
end
