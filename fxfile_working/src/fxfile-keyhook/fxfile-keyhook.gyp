#
# Copyright (c) 2013 Leon Lee author. All rights reserved.
#
#   homepage: http://www.flychk.com
#   e-mail:   mailto:flychk@flychk.com
#
# Use of this source code is governed by a GPLv3 license that can be
# found in the LICENSE file.

{
    'variables':
    {
    },
    
    'includes':
    [
        '../../build/common.gypi',
    ],
    
    'target_defaults':
    {
        'default_configuration' : 'Debug-x86',
        
        'configurations':
        {
            'Debug-x86':
            {
                'inherit_from': ['Debug-x86_Base'],
                
                'msvs_configuration_attributes':
                {
                    'OutputDirectory':       '../../bin',
                    'IntermediateDirectory': '../../obj/fxfile-keyhook/dbg-x86',
                },
                
                'msvs_settings':
                {
                    'VCLinkerTool': 
                    {
                        'OutputFile': '$(OutDir)/$(ProjectName)_dbg.dll',
                        'AdditionalLibraryDirectories':
                        [
                        ],
                        'AdditionalDependencies':
                        [
                        ],
                        'conditions':
                        [
                            [ 'msvs_version == 2008', { 'AdditionalDependencies' : [ '$(INHERIT)', ], }, ],
                            [ 'msvs_version != 2008', { 'AdditionalDependencies' : [ '%(AdditionalDependencies)', ], }, ],
                        ],
                    },
                },
            },
            
            'Release-x86':
            {
                'inherit_from': ['Release-x86_Base'],
                
                'msvs_configuration_attributes':
                {
                    'OutputDirectory':       '../../bin',
                    'IntermediateDirectory': '../../obj/fxfile-keyhook/rel-x86',
                },
                
                'msvs_settings':
                {
                    'VCLinkerTool': 
                    {
                        'OutputFile': '$(OutDir)/$(ProjectName).dll',
                        'AdditionalLibraryDirectories':
                        [
                        ],
                        'AdditionalDependencies':
                        [
                        ],
                        'conditions':
                        [
                            [ 'msvs_version == 2008', { 'AdditionalDependencies' : [ '$(INHERIT)', ], }, ],
                            [ 'msvs_version != 2008', { 'AdditionalDependencies' : [ '%(AdditionalDependencies)', ], }, ],
                        ],
                    },
                },
            },
        },
        'conditions':
        [
            [ 'target_arch!="x86"',
                {
                    'configurations':
                    {
                        'Debug-x64-Unicode':
                        {
                            'inherit_from': ['Debug-x64-Unicode_Base'],
                            'msvs_configuration_attributes':
                            {
                                'OutputDirectory':       '../../bin/x64',
                                'IntermediateDirectory': '../../obj/fxfile-keyhook/dbg-x64',
                                'CharacterSet':          '1',
                            },
                            'defines':
                            [
                                'UNICODE',
                                '_UNICODE',
                            ],
                            'msvs_settings':
                            {
                                'VCLinkerTool': 
                                {
                                    'AdditionalDependencies':
                                    [
                                        'User32.lib',
                                    ],
                                },
                            },
                        },
                        'Release-x64-Unicode':
                        {
                            'inherit_from': ['Release-x64-Unicode_Base'],
                            'msvs_configuration_attributes':
                            {
                                'OutputDirectory':       '../../bin/x64',
                                'IntermediateDirectory': '../../obj/fxfile-keyhook/rel-x64',
                                'CharacterSet':          '1',
                            },
                            'defines':
                            [
                                'UNICODE',
                                '_UNICODE',
                            ],
                            'msvs_settings':
                            {
                                'VCLinkerTool': 
                                {
                                    'AdditionalDependencies':
                                    [
                                        'User32.lib',
                                    ],
                                },
                            },
                        },
                    }
                }
            ]
        ]
    },

    'targets':
    [
        {
            'target_name': 'fxfile-keyhook',
            
            'type': 'shared_library',
            
            'defines':
            [
                'FXFILE_KEYHOOK_EXPORTS',
            ],
            
            'include_dirs':
            [
                './',
            ],
            
            'libraries':
            [
            ],
            
            'msvs_settings':
            {
                'VCLinkerTool': 
                {
                    'AdditionalLibraryDirectories':
                    [
                        '../../bin',
                    ]
                },
            },
            
            'msvs_precompiled_header': 'stdafx.h',
            'msvs_precompiled_source': 'stdafx.cpp',
            

            
            'sources':
            [
                './fxfile-keyhook.cpp',
                './fxfile-keyhook.h',
                './stdafx.cpp',
                './stdafx.h',
            ],
        },
    ],
}