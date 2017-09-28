#!/usr/bin/perl

use v5.10;

use strict;
use warnings;

use Data::Dumper;
use DateTime::Format::Natural;
use Spreadsheet::ParseXLSX;
use Getopt::Long;

my %objects;
GetOptions(
    'xls=s'         => \my $file,
    'ci'            => \$objects{ci},
    'customer'      => \$objects{customer},
    'customer_user' => \$objects{customer_user},
);

if ( !$file || !-f $file ) {
    warn "no XLSX file given";
    exit 1;
}

my @base_cmd = qw(perl /opt/otrs/bin/otrs.Console.pl);

my %data  = _parse_xlsx( $file );
my @order = qw/customer customer_user ci/;

my $requested_imports = grep{ $objects{$_} && $objects{$_} == 1 }@order;
my $import_all        = $requested_imports ? 0 : 1;

say Dumper( \%data );

OBJECT:
for my $object ( @order ) {
    next OBJECT if !$objects{$object} && !$import_all;
    next OBJECT if !$data{$object};

    my $sub = main->can( 'import_' . $object );

    next OBJECT if !$sub;

    $sub->( $data{$object}, \@base_cmd );
}

sub import_ci {
    my ($entities, $base_cmd) = @_;

    my $cmd = 'Admin::ITSM::ConfigItem::Add';
    _run_cmd( $base_cmd, $cmd, $entities, \&_flatten_attribute );
}

sub import_customer_user {
    my ($entities, $base_cmd) = @_;

    my $cmd = 'Admin::CustomerUser::Add';
    _run_cmd( $base_cmd, $cmd, $entities );
}

sub import_customer {
    my ($entities, $base_cmd) = @_;

    my $cmd = 'Admin::CustomerCompany::Add';
    _run_cmd( $base_cmd, $cmd, $entities );
}

sub _run_cmd {
    my ($base_cmd, $cmd, $entities, $sub) = @_;

    for my $entity ( @{ $entities || [] } ) {
        my @args;

        for my $key ( keys %{ $entity || {} } ) {
            my $value = $entity->{$key};

            ($key, $value) = $sub->( $key, $value ) if $sub;

            push @args, '--' . $key, $value;
        }

        say "@{$base_cmd} $cmd @args";
        system @{ $base_cmd }, $cmd, @args;
    }
}

sub _parse_xlsx {
    my ($file) = @_;

    my $parser   = Spreadsheet::ParseXLSX->new;
    my $workbook = $parser->parse($file);

    if ( !defined $workbook ) {
        die $parser->error(), ".\n";
    }

    my %data;

    for my $worksheet ( $workbook->worksheets() ) {

        my $object = $worksheet->get_name;
        my $class;

        if ( $object =~ m{^ci\s*-} ) {
            ($object, $class) = split /\s*-\s*/, $object;
        }

        my ( $row_min, $row_max ) = $worksheet->row_range;
        my ( $col_min, $col_max ) = $worksheet->col_range;

        my @header;
        for my $col ( $col_min .. $col_max ) {
            my $cell = $worksheet->get_cell( $row_min, $col );
            next unless $cell;

            my $header_name = $cell->unformatted;
            $header[$col]   = $header_name;
        }

        for my $row ( $row_min+1 .. $row_max ) {
            my %entity;

            if ( $class ) {
                $entity{class} = $class;
            }

            for my $col ( $col_min .. $col_max ) {

                my $cell = $worksheet->get_cell( $row, $col );
                next unless $cell;

                my $attribute   = $cell->unformatted;
                my $header_name = $header[$col];

                $entity{$header_name} = $attribute;
            }

            push @{ $data{$object} }, \%entity;
        }
    }

    return %data;
}

sub _flatten_attribute {
    my ($key, $value) = @_;

    return ($key, $value) if $key !~ m{\A attr (?:Date (?:Time)? )? -}x;

    my ($type, $attribute) = split /-/, $key;

    if ( -1 < index $type, 'Date' ) {
        my $parser = DateTime::Format::Natural->new;
        my $dt     = $parser->parse_datetime($value);
        $value     = sprintf "%04d-%02d-%02d", $dt->year, $dt->month, $dt->day;

        if ( -1 < index $type, 'Time' ) {
            $value .= sprintf " %02d:%02d:%02d", $dt->hour, $dt->minute, $dt->second;
        }
    }

    return 'attribute', $attribute . '=' . $value;
}
