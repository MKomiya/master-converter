require './xlsx2json'
task :default do
  dirs = Dir.glob('./sheets/*')
  dirs.each do |f|
    c = Xlsx2Json.new
    c.create(f)
    c.run
  end
end
