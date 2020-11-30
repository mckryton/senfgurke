#!/usr/bin/perl

# /usr/bin/perl is the default on Mac OS X, use the following if you
# want to use your PATH to find the best perl available.
# Since a design goal was to use as few 'special' features as possible, this should not be necessary.
# Perl v5.12 and above should be all you need.
#!/usr/bin/env perl

#
# This script generates indexes for directories of BDD feature definitions.
#
# Usage:
#
#   generate_index.pl [options]
#
#   Options:
#
#       # Input:
#           --dir       <dir>       The 'features/' directory to be indexed
#                                   Default: features/
#
#       # Output:
#           --index     <file>      The (base)name of the index file(s) to generate.
#                                   Default: {dir/}index
#
#                                   If the --index has a suffix which matches
#                                   one of the known formats, then the
#                                   appropriate format will be automatically selected.
#
#                                   If --dir was specified more than once, then
#                                   the index file will be written in the
#                                   current directory, otherwise it will be
#                                   written in the directory specified by --dir.
#
#       # Output formats:
#
#           --html      <file>      the name of the html index file to generate
#                                   Default: {index}.html
#
#           --markdown  <file>      the name of the markdown index file to generate
#                                   Default: {index}.md
#
#           --md        <file>      the name of the markdown index file to generate
#                                   Default: {index}.md
#
#   Defaults:
#
#       No index is generated by default.
#       Either the filename specified by --index must have a known format suffix,
#       or at least one of the format-specific options must be provided.
#
#       e.g.:
#
#           % ./generate_index.pl --md
#
#               Scan the features/ directory for .feature files and create a
#               features/index.md markdown file.
#
#           % ./generate_index.pl --index test/feature_map.md
#
#               Scan the features/ directory for .feature files and create a
#               test/feature_map.md markdown file.
#
#           % ./generate_index.pl --dir test/features --index test/feature_map --html --md test/markdown.md
#
#               Scan the features/ directory for .feature files and create a
#               test/feature_map.html and a test/markdown.md file.
#
# The approach taken is:
#
#   Find all *.feature files
#   Parse their 'Feature:' line
#       and any 'description text' which might follow directly after it.
#       Don't parse any Rules, Scenarios, Backgrounds etc.
#           - we just want to know what each feature file is about.
#   Generate one or more index files in the specified format(s).
#

use strict;
use warnings;

use File::Find;     # core module since perl 5
use Getopt::Long;   # core module since perl 5
use Pod::Usage;     # core module since perl 5.6

#
# Currently support index file formats
#
my $format = {
             # Note: each format must define:
             #
             #   getopt:    The getopt option definition, without the :s (optional string) suffix
             #              Aliases for the option may be separated by '|'
             #              (the first option defines the key used in the opt hash)
             #
             #   suffix:    A '|' separated list of possible file name suffixes for the format.
             #              The list is used as a pattern when auto-detecting the
             #              file format from the --index filename.
             #              The first suffix will be used by default when writing files.
             #
             #   generator: The full name of the generator package for the format.
             #              The package MUST provide a new() method for
             #              creating the generator, and a render() method which takes a
             #              feature_tree and returns the text to be written to a file.
             #              TODO: correctly adjust relative links to feature
             #                    files for various --dir and --index combinations
             #              Note: Generators are NOT responsible for writing to the file system!
             #                    Reasons:
             #                      - improved testability.
             #                      - the ability to 'stream' output, if the need should arise.
             html     => { getopt => 'html',        suffix => 'html|htm',    generator => 'generator::html'     },
             json     => { getopt => 'json',        suffix => 'json',        generator => 'generator::json'     },
             markdown => { getopt => 'markdown|md', suffix => 'md|markdown', generator => 'generator::markdown' },
             yaml     => { getopt => 'yaml|yml',    suffix => 'yaml|yml',    generator => 'generator::yaml'     },
             };

#
# MAIN
#

# "if( not caller )" Trick to help testing:
#   Only *run* the main code if this script was called directly from the command line.
#   When this script is loaded by the test/*.t files, caller will be 'true' and
#   thus we'll just be left with a buch of subs and packages
if( not caller )
    {

    my $opt = {};
    GetOptions( $opt,
                'verbose',          # turn on verbose output
                'dir=s',            # directory to scan for feature files
                'index=s',          # (base)name of the index file(s) to generate
                # dynamically add options for each known format, e.g.: --html, --markdown etc
                map { "$format->{$_}{getopt}:s" } keys %{$format},
                )
        or pod2usage( -message => 'oh dear... help still needs to be implemented :-(' );

    derive_default_options( $opt );

    # get a tree of feature information to put into the index
    my $feature_tree = scan_feature_dirs( search => $opt->{dir} );

    my $generated = 0;
    foreach my $fmt ( sort keys %{$format} )
        {
        if( $opt->{$fmt} )
            {
            $opt->{verbose} and printf "generating %s index => %s ...\n", $fmt, $opt->{$fmt};
            my $g = $format->{$fmt}{generator}->new();
            my $t = $g->render( $feature_tree );
            if( open( my $fh, ">", $opt->{$fmt} ) )
                {
                print $fh $t;
                close $fh;
                $generated++;
                }
            else
                {
                printf STDERR "Error: failed to write to %s: %s\n", $opt->{$fmt}, $!;
                }
            }
        }

    if( not $generated )
        {
        printf STDERR "Error: No index files were generated.\n";    # TODO: 2020-11-16 list available formats
        exit 1;
        }

    exit;
    }

sub derive_default_options
    {
    my $opt = shift;

    # --dir and --index are the only two independent options that can get
    # static defaults.
    $opt->{dir}   = 'features' if not $opt->{dir};
    $opt->{index} = 'index'    if not $opt->{index};

    # remove trailing /'s from search directory for cosmetic reasons
    $opt->{dir} =~ s{/*$}{};

    # if --index has a known suffix, make sure that that type of file will be generated.
    if( $opt->{index} =~ /(?<base>.+?)\.(?<suffix>\w+)$/ )
        {
        my $index_base   = $+{base};
        my $index_suffix = $+{suffix};
        SUFFIX_MATCH:
        foreach my $fmt ( keys %{$format} )
            {
            # re-use the suffix patterns as regex to match against known suffixes
            my $suffix_re = $format->{$fmt}{suffix};
            if( $index_suffix =~ /$suffix_re/ )
                {
                $opt->{$fmt}  //= $opt->{index};
                $opt->{index}  = $index_base;
                last SUFFIX_MATCH;
                }
            }
        }

    # setup the format => filename mappings for all requested formats
    foreach my $fmt ( keys %{$format} )
        {
        next if not defined $opt->{$fmt};
        # use the suffix patterns as regex to match against known suffixes
        my $suffix_re      = $format->{$fmt}{suffix};
        my $default_suffix = $format->{$fmt}{suffix} =~ s{\|.*}{}r;
        if( not $opt->{$fmt}                     ) { $opt->{$fmt}  = $opt->{index};                              }
        if( not $opt->{$fmt} =~ m:\.$suffix_re$: ) { $opt->{$fmt} .= sprintf ".%s",   $default_suffix;           }
        if( not $opt->{$fmt} =~ m:/:             ) { $opt->{$fmt}  = sprintf "%s/%s", $opt->{dir}, $opt->{$fmt}; }
        }
    }

sub scan_feature_dirs
    {
    my $param        = { @_ };
    my $search       = $param->{search};
    my $node         = {
                       name     => '',
                       type     => 'dir',
                       contents => [],
                       count    => 0,
                       parent   => undef,
                       };
    my $feature_tree = $node;
    my $find_opts    = {
                       no_chdir     => 1,   # don't chdir into each subdirectory - keep all paths relative to topdir
                       follow       => 0,   # must be 0 to use preprocess
                       follow_fast  => 0,   # must be 0 to use preprocess
                       preprocess   =>  sub {
                                            my $leaf = $File::Find::dir =~ s{.*/}{}r;
                                            push @{$node->{contents}}, {
                                                                       name     => $leaf
                                                                                    =~ s{_+}{ }gr        # replace _'s with spaces
                                                                                    =~ s{\b\w}{\u$&}gr,  # uppercase each first character
                                                                       type     => 'dir',
                                                                       contents => [],
                                                                       count    => 0,
                                                                       parent   => $node,
                                                                       };
                                            $node = $node->{contents}[-1];
                                            # return list of files/directories we were given
                                            return @_;
                                            },
                       wanted       =>  sub {
                                            return if not $File::Find::name =~ /\.feature$/;
                                            if( my $feature = parse_feature( $File::Find::name, $search ) )
                                                {
                                                push @{$node->{contents}}, $feature;
                                                my $parent = $node;
                                                while( $parent )
                                                    {
                                                    $parent->{count}++;
                                                    $parent = $parent->{parent};
                                                    }
                                                }
                                            },
                       postprocess  =>  sub {
                                            my $parent = $node->{parent};

                                            # discard empty directories
                                            pop @{$parent->{contents}} if not $node->{count};

                                            # sort contents by the name displayed in the index
                                            @{$node->{contents}} = map { $_->[0] }
                                                                   sort { $a->[1] cmp $b->[1] }
                                                                   map { [ $_, lc( $_->{name} ) ] }
                                                                   @{$node->{contents}};

                                            # remove scaffolding
                                            delete $node->{count};
                                            delete $node->{parent};

                                            # move up to parent directory
                                            $node = $parent;
                                            },
                       };

    find( $find_opts, ( $search ) );

    # remove the 'top nodes' that only have a single sub-directory in their contents
    while( @{$feature_tree->{contents}} == 1 and $feature_tree->{contents}[0]{type} eq 'dir' )
        {
        $feature_tree = $feature_tree->{contents}[0];
        }

    return $feature_tree;
    }

sub parse_feature
    {
    my $file        = shift;
    my $topdir      = shift;

    my $description = [];
    my $info        = {
                      name => undef,
                      type => 'feature',
                      file => $file =~ s{^$topdir/+}{}r,
                      desc => $description,
                      };

    my $ignore_line_re  = qr/^\s*[#@]/;
    my $keyword_line_re = qr/^\s*(?<keyword>[^:]*[^:\s]+[^:]*?)\s*:\s*(?<name>\S.*?)\s*$/;

    if( open( my $fd, '<', $file ) )
        {
        my $block      = 0;
        my $min_indent = 1000;  # +infinity when it comes to indentation of feature files
        while( my $line = <$fd> )
            {
            next        if $line  =~ /$ignore_line_re/;     # skip all comments and tag lines
            $block++    if $line  =~ /$keyword_line_re/;    # have we started a new block?
            last        if $block  > 1;                     # stop after finishing the 1st block
            next        if $block == 0;                     # stop after finishing the 1st block

            my $name = $+{name};    # grab name from keyword match above (most recent regex)

            # chomp regardless of NL, CRNL style and expand all tabs to spaces
            $line = expandtabs( $line =~ s{\v+}{}r );

            if( $name )
                {
                # The first line of the first block will be 'Feature:',
                # 'Ability:', 'Business Need:' or any other language specific synonym.
                # BUG:  Since we're not actually checking for *known* keywords,
                #       any line containing a : will be considered the start of
                #       a new block.
                #       The regex above makes sure that there is at least 1
                #       non-: and 1 non-space before and after each keyword.
                $info->{name} = ucfirst( $name =~ s{\s+}{ }gr ); # squash tabs and spaces in the name
                }
            elsif( $line =~ /^\s*$/ )
                {
                # squash leading and sequential empty lines
                push @{$description}, ''    if @{$description} and $description->[-1] ne '';
                }
            else
                {
                # keep any text within the block
                push @{$description}, $line;
                my $indent = length( $line =~ s{^(\s+).*}{$1}r );
                $min_indent = $indent if $indent < $min_indent;
                }
            }

        # cleanup possible trailing blank line
        pop @{$description}    if @{$description} and $description->[-1] eq '';

        # remove common indent from all lines
        my $strip_re = qr{^\s{$min_indent}};
        foreach my $line ( @{$description} )
            {
            $line =~ s{$strip_re}{};
            }

        close $fd;
        }
    else
        {
        printf STDERR "Error: failed to open '$file': $!\n";
        }

    return $info;
    }

sub expandtabs
    {
    my $text = shift;
    while( $text and ( my $t = index( $text, "\t" ) ) >= 0 )
        {
        my $r = ( 8 - ( $t % 8 ) ) || 8;    # brackets to help YOU understand what I indended :-)
        $text =~ s{\t}{' ' x $r}e;
        }
    return $text;
    }

####################################################################################################
#
# HTML FORMATTER
#

package generator::html;

use strict;
use warnings;

sub new { return bless {}, shift; }

sub render
    {
    my $self         = shift;
    my $feature_tree = shift;  # an AST as generated by scan_feature_dirs()

    my $index_html   = $self->feature_tree_as_html( $feature_tree );

    # set up the dynamic contents to be inserted into the HTML template
    my $html_var     = {
                       feature_index => join( "\n", @{$index_html} ),
                       };

    my $html = $self->html_template();
    $html =~ s{
              {(?<name>\w+)}                # replace placeholders in the html template...
              }
              {
              exists $html_var->{$+{name}}
                  ? $html_var->{$+{name}}   # with their defined values
                  : $&                      # or leave them unchanged
              }egx;

    return $html;
    }

sub feature_tree_as_html
    {
    my $self  = shift;
    my $node  = shift;
    my $html  = shift || [];

    if( $node->{type} eq 'dir' )
        {
        push @{$html}, '<div class=dir>';
        push @{$html}, sprintf "<label>%s</label>", $node->{name}    if $node->{name};
        push @{$html}, '<ul>';
        foreach my $child ( @{$node->{contents}} )
            {
            push @{$html}, '<li>';
            $self->feature_tree_as_html( $child, $html ); # recursion - yay!
            push @{$html}, '</li>';
            }
        push @{$html}, '</ul>';
        push @{$html}, '</div>';
        }
    else # only one other option: feature!
        {
        push @{$html}, '<div class=feature>';
        push @{$html}, '<label>';
        push @{$html}, sprintf "<a href=\"%s\">%s</a>", $node->{file}, $node->{name};
        push @{$html}, '</label>';
        if( @{$node->{desc}} )
            {
            push @{$html}, '<pre class=description><code>';
            push @{$html}, @{$node->{desc}};
            push @{$html}, '</code></pre>';
            }
        else
            {
            push @{$html}, '<div class="empty">no description</div>';
            }
        push @{$html}, '</div>';
        }

    return $html;
    }

sub html_template
    {
    return <<_EO_TEMPLATE_;
<!doctype html>
<html>
<head>
    <title> Features </title>
    <meta charset = "utf-8">

    <style>

        ul                              {
                                        list-style:     none;
                                        border-left:    1px solid #ddd;
                                        padding-left:   2em;
                                        }

        .dir label                      {
                                        font-weight:    bold;
                                        line-height:     2em;
                                        }

        .feature label                  {
                                        font-weight:    bold;
                                        }

        .feature .description           {
                                        display:        none;
                                        padding:        0 1em 1em;  /* top, right, bottom( , left ) - clockwise, copy missing from opposite side */
                                        border:         1px solid blue;
                                        border-radius:  1ch;
                                        background:     #eef;
                                        }

        .feature .empty                 {
                                        display:        none;
                                        padding:        1em;
                                        border:         1px solid #aaa;
                                        border-radius:  1ch;
                                        background:     #fafafa;
                                        color:          #888;
                                        font-style:     italic;
                                        }

        .feature:hover .description,
        .feature:hover .empty           {
                                        display:        block;
                                        }

    </style>

</head>

<body>

    <main>

        <div id=index>
            {feature_index}
        </div>

    </main>

</body>

</html>
_EO_TEMPLATE_
    }

####################################################################################################
#
# MARKDOWN FORMATTER
#
# References:
#
#       https://daringfireball.net/projects/markdown/
#       https://tools.ietf.org/html/rfc7763
#       https://tools.ietf.org/html/rfc7764
#       https://help.github.com/articles/github-flavored-markdown/
#       https://github.com/github/markup/tree/master#html-sanitization
#

package generator::markdown;

use strict;
use warnings;

sub new { return bless {}, shift; }

sub render
    {
    my $self            = shift;
    my $feature_tree    = shift;  # an AST as generated by scan_feature_dirs()

    my $markdown        = $self->feature_tree_as_markdown( $feature_tree );

    return join( "\n", @{$markdown} );
    }

sub feature_tree_as_markdown
    {
    my $self     = shift;
    my $node     = shift;
    my $markdown = shift || [];
    my $depth    = ( shift || 0 ) + 1;

    if( $node->{type} eq 'dir' )
        {
        push @{$markdown}, sprintf "%s %s", '#' x $depth, $node->{name};
        push @{$markdown}, '';
        foreach my $child ( @{$node->{contents}} )
            {
            $self->feature_tree_as_markdown( $child, $markdown, $depth ); # recursion - yay!
            }
        }
    else # only one other option: feature!
        {
        push @{$markdown}, sprintf "* [%s](%s)", $node->{name}, $node->{file};
        push @{$markdown}, '';
        if( @{$node->{desc}} )
            {
            # TODO: 2020-11-17 Problem with markdown: anything indented > 4 spaces is considered 'code'
            # The descriptions only have their relative indentation - which can
            # give us surprising combinations of 'text' and 'code'.
            # Alternative: remove all indentation?
            #
            # plain text => strange indentation
            # push @{$markdown}, @{$node->{desc}};
            #
            # common indent => bizarre 'hash' errors - but only via Markdown.pl (John Gruber, 2004)
            # push @{$markdown}, map { s{^\s*}{  }r } @{$node->{desc}};
            #
            # ensure that all lines with text are indented 2 spaces, to line up
            # with the feature name, while keeping empty lines empty
            push @{$markdown}, map { $_ ? s{^\s*}{  }r : '' } @{$node->{desc}};
            }
        else
            {
            push @{$markdown}, '  _no description_';
            }
        push @{$markdown}, '';
        }

    return $markdown;
    }

####################################################################################################
#
# YAML FORMATTER
#
# Note: the YAML module is NOT in perl's core distribution
# For this purpose, we don't really need it, just a simple recursive structure
#

package generator::yaml;

use strict;
use warnings;

sub new { return bless {}, shift; }

sub render
    {
    my $self         = shift;
    my $feature_tree = shift;  # an AST as generated by scan_feature_dirs()

    my $yaml         = $self->feature_tree_as_yaml( $feature_tree );

    return sprintf "---\n%s\n", join( "\n", @{$yaml} );
    }

sub feature_tree_as_yaml
    {
    my $self   = shift;
    my $node   = shift;
    my $yaml   = shift || [];
    my $depth  = ( shift || 0 ) + 1;

    my $indent = '    ' x $depth;

    if( $node->{type} eq 'dir' )
        {
        push @{$yaml}, sprintf "%s- name: %s", $indent, yaml_string( $node->{name} );
        push @{$yaml}, sprintf "%s  contents:", $indent;
        foreach my $child ( @{$node->{contents}} )
            {
            $self->feature_tree_as_yaml( $child, $yaml, $depth ); # recursion - yay!
            }
        }
    else # only one other option: feature!
        {
        push @{$yaml}, sprintf "%s- name: %s", $indent, yaml_string( $node->{name} );
        push @{$yaml}, sprintf "%s  file: %s", $indent, yaml_string( $node->{file} );
        if( @{$node->{desc}} )
            {
            push @{$yaml}, sprintf "%s  desc:", $indent;
            foreach my $line ( @{$node->{desc}} )
                {
                push @{$yaml}, sprintf "%s      - %s", $indent, yaml_string( $line );
                }
            }
        else
            {
            push @{$yaml}, sprintf "%s  desc: []", $indent;
            }
        }

    return $yaml;
    }

sub yaml_string
    {
    my $string = shift;
    return $string =~ /^\s|\s$|^$/
            ? $string =~ /'/
                ? sprintf( "\"%s\"", $string =~ s/"/\\"/gr )
                : sprintf( "\'%s\'", $string )
            : $string;
    }

####################################################################################################
#
# JSON FORMATTER
#
# Note: the JSON module is NOT in perl's core distribution
# For this purpose, we don't really need it, just a simple recursive structure
#

package generator::json;

use strict;
use warnings;

sub new { return bless {}, shift; }

sub render
    {
    my $self         = shift;
    my $feature_tree = shift;  # an AST as generated by scan_feature_dirs()

    my $json = $self->feature_tree_as_json( $feature_tree );
    $json->[-1] =~ s/,$//;

    return sprintf "%s\n", join( "\n", @{$json} );
    }

sub feature_tree_as_json
    {
    my $self   = shift;
    my $node   = shift;
    my $json   = shift || [];
    my $depth  = ( shift || 0 ) + 1;

    my $indent = '  ' x ( $depth - 1 );

    if( $node->{type} eq 'dir' )
        {
        push @{$json}, sprintf "%s{",                 $indent;
        push @{$json}, sprintf "%s\"name\":     %s,", $indent, json_string( $node->{name} );
        push @{$json}, sprintf "%s\"contents\": [",   $indent;
        foreach my $child ( @{$node->{contents}} )
            {
            $self->feature_tree_as_json( $child, $json, $depth + 6 ); # recursion - yay!
            }
        $json->[-1] =~ s/,$//;
        push @{$json}, sprintf "%s          ]", $indent;
        push @{$json}, sprintf "%s},",          $indent;
        }
    else # only one other option: feature!
        {
        push @{$json}, sprintf "%s{",             $indent;
        push @{$json}, sprintf "%s\"name\": %s,", $indent, json_string( $node->{name} );
        push @{$json}, sprintf "%s\"file\": %s,", $indent, json_string( $node->{file} );
        if( @{$node->{desc}} )
            {
            push @{$json}, sprintf "%s\"desc\": [", $indent;
            foreach my $line ( @{$node->{desc}} )
                {
                push @{$json}, sprintf "%s      %s,", $indent, json_string( $line );
                }
            $json->[-1] =~ s/,$//;
            push @{$json}, sprintf "%s      ]", $indent;
            }
        else
            {
            push @{$json}, sprintf "%s\"desc\": []", $indent;
            }
        push @{$json}, sprintf "%s},", $indent;
        }

    return $json;
    }

sub json_string
    {
    my $string = shift;
    return sprintf( "\"%s\"", $string =~ s/"/\\"/gr )
    }

# keep 'require' happy when this script is 'imported' while running tests
1;