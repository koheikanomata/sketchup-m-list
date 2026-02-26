# frozen_string_literal: true
# M List - エンティティ関連

module EstimateAuto
  def self.entity_type_label(e)
    case e
    when Sketchup::Group then "グループ"
    when Sketchup::ComponentInstance then "コンポーネント"
    when Sketchup::Face then "面"
    when Sketchup::Edge then "線"
    else "その他"
    end
  end

  def self.entity_tag(e)
    return "" unless e.respond_to?(:layer)
    layer = e.layer
    return "" unless layer
    name = (layer.respond_to?(:display_name) ? layer.display_name : nil) || layer.name.to_s
    name = "" if name.to_s.strip.empty?
    (name == "Layer0" || name.downcase == "layer 0") ? "Untagged" : name
  end

  def self.entity_name(e)
    case e
    when Sketchup::ComponentInstance
      name = e.name.to_s
      name = e.definition.name.to_s if name.empty?
      name.empty? ? "(無名コンポーネント)" : name
    when Sketchup::Group
      name = e.name.to_s
      name.empty? ? "Group" : name
    when Sketchup::Face, Sketchup::Edge
      nil
    else
      ""
    end
  end

  def self.size_string(e)
    return nil unless e.respond_to?(:bounds)
    bb = e.bounds
    w = bb.width.to_mm.round
    h = bb.height.to_mm.round
    d = bb.depth.to_mm.round
    "#{w} x #{h} x #{d} mm"
  end

  def self.get_selected_objects
    return [] unless model
    model.selection.select { |e| e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance) }
  end
end
