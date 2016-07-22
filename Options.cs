using System;
using System.Collections.Generic;
using System.Text;

class Options
{
  public const string SPLIT_CHAR = ":";

  public static readonly char[] SPLIT_CHAR_A = SPLIT_CHAR.ToCharArray();

  private Dictionary<AllowedOptions, string> _options;
  private AllowedOptions? _greedyOption;
  
  //enum Helpers
  private Dictionary<string, AllowedOptions> _s2ao;
  private Dictionary<AllowedOptions, string> _ao2s;
  private List<AllowedOptions> _activeOptions;

  private void buildEnumHelpers()
  {
    _s2ao = new Dictionary<string, AllowedOptions>();
    _ao2s = new Dictionary<AllowedOptions, string>();

    foreach (AllowedOptions ao in Enum.GetValues(typeof(AllowedOptions)))
    {
      _s2ao.Add(ao.ToString(), ao);
      _ao2s.Add(ao, ao.ToString());
    }
  }

  public Options(string[] args, AllowedOptions? GreedyOption)
  {
    _greedyOption = GreedyOption;
    _options = new Dictionary<AllowedOptions, string>();
    _activeOptions = new List<AllowedOptions>();
    buildEnumHelpers();
    _options = parse(args);
    _activeOptions = new List<AllowedOptions>(_options.Keys);
  }

  public string this[AllowedOptions Option] 
  {
    get
    {
      if (_options.ContainsKey(Option)) return _options[Option];
      return "";
    }
    set
    {
      if (_options.ContainsKey(Option))
        _options[Option] = value;
    }
  }

  public string this[string Option]
  {
    get
    {
      if (_s2ao.ContainsKey(Option))
      {
        AllowedOptions o = _s2ao[Option];
        if(_activeOptions.Contains(o))
          return _options[_s2ao[Option]];
      }
      return "";
    }
  }

  private Dictionary<AllowedOptions,string> parse(string[] args)
  {
    Dictionary<AllowedOptions, string> result = new Dictionary<AllowedOptions, string>();

    //establish a cursor and cl in case their is a greedy option 
    int cursor = Environment.GetCommandLineArgs()[0].Length; //start checking after the program name
    string cmdline = Environment.CommandLine;

    string goName = "";
    if (HasGreedyOption) goName = _ao2s[(AllowedOptions)_greedyOption];
    
    foreach (string sArg in args) {
     
      string[] parts = sArg.Split(SPLIT_CHAR_A, 2, StringSplitOptions.RemoveEmptyEntries);

      //get the key removing whitespace
      string key = parts[0].Trim();

      if (_s2ao.ContainsKey(key))
      {
        
        cursor = cmdline.IndexOf(key, cursor) + key.Length + 1;

        string value = "";

        if (parts.Length == 2) value = parts[1];

        if (HasGreedyOption && goName == key) value = cmdline.Substring(cursor);

        result.Add(_s2ao[key], value);

        if (HasGreedyOption && goName == key) break; //quit loop if we dealt with the greedy option

      }
      else
      {
        Console.WriteLine("* Invalid option: " + key);
        Environment.Exit(-1);
      }

    }

    return result;

    //end Parse
  }

  public bool OptionExists(AllowedOptions Option) {
    return _options.ContainsKey(Option);
  }

  public AllowedOptions[] ActiveOptions
  {
    get { return _activeOptions.ToArray(); }
  }

  public bool HasGreedyOption { get { return _greedyOption != null; } }

  public AllowedOptions? GreedyOption {  get { return _greedyOption; } }

  public override string ToString()
  {
    StringBuilder result = new StringBuilder("Active Options: \r\n");
    foreach(AllowedOptions o in _ao2s.Keys)
      if(_options.ContainsKey(o))
      result.AppendFormat(" - {0}: {1}\r\n", _ao2s[o], _options[o]);
    return result.ToString();
  }

  //end class Options
}
